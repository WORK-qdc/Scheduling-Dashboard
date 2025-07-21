// sort state
let sortColumn   = null;       // e.g. 'title', 'start', 'AssignedEmployee', ...
let sortDirection = 'asc';     // or 'desc'

function showLoading() {
  document.getElementById('loading-bar').classList.add('active');
}

function hideLoading() {
  setTimeout(() => {
    document.getElementById('loading-bar').classList.remove('active');
  }, 300);
}

document.addEventListener('DOMContentLoaded', async () => {
  const listBtn           = document.getElementById('list-view-btn');
  const calBtn            = document.getElementById('calendar-view-btn');
  const listContainer     = document.getElementById('list-container');
  const calendarContainer = document.getElementById('calendar-container');
  let isCollapsedMode     = true; // Default to collapsed

  const msalConfig = {
    auth: {
      clientId: '8e23d112-104b-4e6a-a57d-7e3a2a61d837',
      authority: 'https://login.microsoftonline.com/20e27700-b670-4553-a27c-d8e2583b3289',
      redirectUri: 'https://work-qdc.github.io/Scheduling-Dashboard/'
    },
    cache: { cacheLocation: 'localStorage' }
  };

  document.getElementById('collapsed-mode-toggle').addEventListener('click', () => {
    isCollapsedMode = !isCollapsedMode;
    const toggle = document.getElementById('collapsed-mode-toggle');
    toggle.textContent = isCollapsedMode ? 'Expanded View' : 'Collapsed View';
    renderList();
  });
function openEditModal(entry) {
  document.getElementById('edit-entry-id').value            = entry.id;
  document.getElementById('edit-title').value               = entry.title;
  document.getElementById('edit-topic').value               = entry.topic;
  document.getElementById('edit-status').value              = entry.status;
  document.getElementById('edit-start').value               = formatDateOnly(entry.start);
  document.getElementById('edit-end').value                 = formatDateOnly(entry.end);
  document.getElementById('edit-location').value            = entry.location;
  document.getElementById('edit-link').value                = entry.link;
  document.getElementById('edit-industry').value            = entry.industry;
  document.getElementById('edit-desc').value                = entry.desc;
  document.getElementById('edit-applicationdeadline').value = formatDateOnly(entry.applicationdeadline);
  document.getElementById('edit-type').value                = entry.internalExternal;
  document.getElementById('edit-AssignedEmployee').value    = entry.AssignedEmployee;
  document.getElementById('edit-notes').value               = entry.notes;

  new bootstrap.Modal(document.getElementById('editEntryModal')).show();
}
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
    const site = await g.api(`/sites/tcco.sharepoint.com:/sites/SchedulingToolTest:/`).get();
    return site.id;
  }

  async function initLists() {
    showLoading();
    await signIn();
    const token = await getToken(['Sites.Read.All']);
    const g     = client(token);
    const sid   = await getSiteId(g);
    const res   = await g.api(`/sites/${sid}/lists`).get();
    EMP_LIST_ID   = res.value.find(l => l.displayName === 'SchedulingEmployeeTest').id;
    SCHED_LIST_ID = res.value.find(l => l.displayName === 'TestTableScheduling').id;
    hideLoading();
  }

  async function fetchEmployees() {
    showLoading();
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
    populateEmployeeSelects();
    renderEmployeeTable();
    hideLoading();
  }

  async function fetchEntries() {
    showLoading();
    const token = await getToken(['Sites.Read.All']);
    const g     = client(token);
    const sid   = await getSiteId(g);
    const r     = await g
      .api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items`)
      .expand('fields')
      .get();
    window.entries = r.value.map(i => ({
      id:                  i.id,
      title:               i.fields.Title || '',
      topic:               i.fields.Topic || '',
      status:              i.fields.RegistrationStatus || '',
      start:               i.fields.StartDate || '',
      end:                 i.fields.EndDate || '',
      location:            i.fields.Location || '',
      link:                i.fields.Link || '',
      industry:            i.fields.Industry || '',
      desc:                i.fields.Description || '',
      applicationdeadline: i.fields.ApplicationDeadline || '',
      internalExternal:    i.fields.Internal_x002f_External || '',
      AssignedEmployee:    i.fields.AssignedEmployee || '',
      notes:               i.fields.Notes || ''
    }));
    renderList();
    hideLoading();
  }

  async function updateEntry() {
    try {
      showLoading();
      const id = document.getElementById('edit-entry-id').value;
      const fields = {
        Title:                 document.getElementById('edit-title').value,
        Topic:                 document.getElementById('edit-topic').value,
        RegistrationStatus:    document.getElementById('edit-status').value,
        StartDate:             document.getElementById('edit-start').value,
        EndDate:               document.getElementById('edit-end').value,
        Location:              document.getElementById('edit-location').value,
        Link:                  document.getElementById('edit-link').value,
        Industry:              document.getElementById('edit-industry').value,
        Description:           document.getElementById('edit-desc').value,
        ApplicationDeadline:   document.getElementById('edit-applicationdeadline').value,
        Internal_x002f_External: document.getElementById('edit-type').value,
        AssignedEmployee:      document.getElementById('edit-AssignedEmployee').value,
        Notes:                 document.getElementById('edit-notes').value
      };

      const token = await getToken(['Sites.ReadWrite.All']);
      const g     = client(token);
      const sid   = await getSiteId(g);

      await g
        .api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items/${id}`)
        .patch({ fields });

      await fetchEntries();
      bootstrap.Modal
        .getInstance(document.getElementById('editEntryModal'))
        .hide();
      clearEditForm();
    } catch (err) {
      console.error('updateEntry failed:', err);
      alert(`Error updating entry:\n${err.message}`);
    } finally {
      hideLoading();
    }
  }

  async function addEmployee() {
    const name  = prompt('Enter employee name:');
    const email = prompt('Enter employee email:');
    if (!name || !email) return;
    showLoading();
    const token = await getToken(['Sites.ReadWrite.All']);
    const g     = client(token);
    const sid   = await getSiteId(g);
    await g.api(`/sites/${sid}/lists/${EMP_LIST_ID}/items`).post({
      fields: { Title: name, Employeesemail: email }
    });
    await fetchEmployees();
  }

  async function addEntry() {
    try {
      showLoading();
      const token = await getToken(['Sites.ReadWrite.All']);
      const g     = client(token);
      const sid   = await getSiteId(g);

      const sel                   = document.getElementById('new-AssignedEmployee');
      const assignedEmployeeName  = sel.value;
      const assignedEmployeeEmail = (window.employees.find(e => e.name === assignedEmployeeName) || {}).email;

      const fields = {
        Title:                 document.getElementById('new-title').value,
        Topic:                 document.getElementById('new-topic').value,
        RegistrationStatus:    document.getElementById('new-status').value,
        TypeofEngagement:      document.getElementById('new-type').value,
        StartDate:             document.getElementById('new-start').value,
        EndDate:               document.getElementById('new-end').value,
        Location:              document.getElementById('new-location').value,
        Link:                  document.getElementById('new-link').value,
        Industry:              document.getElementById('new-industry').value,
        Description:           document.getElementById('new-desc').value,
        ApplicationDeadline:   document.getElementById('new-applicationdeadline').value,
        Internal_x002f_External: document.getElementById('new-type').value,
        AssignedEmployee:      assignedEmployeeName,
        Notes:                 document.getElementById('new-notes').value
      };

      await g.api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items`).post({ fields });

      // Send calendar invite if email valid...
      if (assignedEmployeeEmail && assignedEmployeeEmail.includes('@')) {
        const event = {
          subject: fields.Title || 'Scheduled Event',
          body: {
            contentType: 'Text',
            content:     fields.Description || 'Scheduled event via scheduling tool.'
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
          attendees: [{
            emailAddress: {
              address: assignedEmployeeEmail,
              name:    assignedEmployeeName
            },
            type: 'required'
          }],
          responseRequested: true,
          isOnlineMeeting:   false
        };

        try {
          await g
            .api('/me/events')
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

  // expose addEntry globally
  window.addEntry = addEntry;

  function clearAddForm() {
    document.getElementById('new-title').value                       = '';
    document.getElementById('new-topic').value                       = '';
    document.getElementById('new-status').value                      = 'Pending';
    document.getElementById('new-start').value                       = '';
    document.getElementById('new-end').value                         = '';
    document.getElementById('new-location').value                    = '';
    document.getElementById('new-link').value                        = '';
    document.getElementById('new-industry').value                    = '';
    document.getElementById('new-desc').value                        = '';
    document.getElementById('new-applicationdeadline').value         = '';
    document.getElementById('new-AssignedEmployee').value            = '';
    document.getElementById('new-notes').value                       = '';
  }

  function populateEmployeeSelects() {
    const sel       = document.getElementById('new-AssignedEmployee');
    const editSel   = document.getElementById('edit-AssignedEmployee');
    const filterSel = document.getElementById('employee-filter');
    const calSel    = document.getElementById('calendar-employee-filter');

    if (sel)       sel.innerHTML       = '<option value="">Select Speaker</option>';
    if (editSel)   editSel.innerHTML   = '<option value="">Select Speaker</option>';
    if (filterSel) filterSel.innerHTML = '<option value="">All Speakers</option>';
    if (calSel)    calSel.innerHTML    = '<option value="">All Speakers</option>';

    window.employees.forEach(e => {
      if (sel) {
        const o1 = document.createElement('option');
        o1.value       = e.name;
        o1.textContent = e.name;
        sel.appendChild(o1);
      }
      if (editSel) {
        const o2 = document.createElement('option');
        o2.value       = e.name;
        o2.textContent = e.name;
        editSel.appendChild(o2);
      }
      if (filterSel) {
        const o3 = document.createElement('option');
        o3.value       = e.name;
        o3.textContent = e.name;
        filterSel.appendChild(o3);
      }
      if (calSel) {
        const o4 = document.createElement('option');
        o4.value       = e.name;
        o4.textContent = e.name;
        calSel.appendChild(o4);
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
        <td>
          <button class="btn btn-sm btn-danger delete-emp">Delete</button>
        </td>
      `;
      tr.querySelector('.delete-emp').addEventListener('click', async () => {
        if (!confirm(`Delete employee ${e.name}?`)) return;
        showLoading();
        try {
          const token = await getToken(['Sites.ReadWrite.All']);
          const g     = client(token);
          const sid   = await getSiteId(g);
          await g
            .api(`/sites/${sid}/lists/${EMP_LIST_ID}/items/${e.id}`)
            .delete();
          await fetchEmployees();
        } catch (err) {
          console.error(err);
          alert('Error deleting employee:\n' + err.message);
        } finally {
          hideLoading();
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
    return `<span class="status-badge ${cls}">${status || 'Pending'}</span>`;
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

    const selectedEmployee = document.getElementById('calendar-employee-filter').value;
    const filteredEvents = window.entries
      .filter(e => !selectedEmployee || e.AssignedEmployee === selectedEmployee)
      .map(e => ({
        id:              e.id,
        title:           e.title || '(No Title)',
        start:           e.start,
        end:             e.end,
        backgroundColor: e.status === 'Confirmed'
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
          topic:    e.topic,
          speaker:  e.AssignedEmployee
        }
      }));

    const calendar = new FullCalendar.Calendar(calendarEl, {
      initialView:   'dayGridMonth',
      height:        '100%',
      contentHeight: 'auto',
      headerToolbar: {
        left:   'prev,next today',
        center: 'title',
        right:  'dayGridMonth,timeGridWeek,listWeek'
      },
      events: filteredEvents,
      eventClick(info) {
        const { title, start, end, extendedProps, id } = info.event;
        const modal = getOrCreateEventModal();

        document.getElementById('eventModalBody').innerHTML = `
          <p><strong>Title:</strong>    ${title}</p>
          <p><strong>Status:</strong>   ${getStatusBadge(extendedProps.status)}</p>
          <p><strong>Start:</strong>    ${formatDate(start)}</p>
          <p><strong>End:</strong>      ${formatDate(end)}</p>
          <p><strong>Location:</strong> ${extendedProps.location || 'N/A'}</p>
          <p><strong>Speaker:</strong>  ${extendedProps.speaker || '<em>Unassigned</em>'}</p>
        `;

        modal.querySelector('#editEventBtn').onclick = () => {
  const entry = window.entries.find(e => e.id == id);
  if (entry) openEditModal(entry);
};
        modal.querySelector('#deleteEventBtn').onclick = async () => {
          if (!confirm(`Delete "${title}"?`)) return;
          bootstrap.Modal.getInstance(modal).hide();
          showLoading();
          try {
            const token = await getToken(['Sites.ReadWrite.All']);
            const g     = client(token);
            const sid   = await getSiteId(g);
            await g
              .api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items/${id}`)
              .delete();
            await fetchEntries();
          } catch (err) {
            console.error('deleteEntry failed:', err);
            alert('Error deleting entry: ' + err.message);
          } finally {
            hideLoading();
          }
        };

        modal.dataset.entryId    = id;
        modal.dataset.eventData  = JSON.stringify({
          title,
          status: extendedProps.status,
          start,
          end,
          location: extendedProps.location,
          speaker:  extendedProps.speaker
        });

        new bootstrap.Modal(modal).show();
      }
    });

    calendar.render();
  }

  // === UPDATED createEventModal with footer Edit/Delete buttons ===
  function createEventModal() {
    const modal = document.createElement('div');
    modal.className = 'modal fade';
    modal.id        = 'eventModal';

    modal.innerHTML = `
      <div class="modal-dialog">
        <div class="modal-content">
          <button type="button"
                  class="btn btn-sm btn-light border toggle-fullscreen"
                  title="Toggle Fullscreen">
            <i class="bi bi-arrows-fullscreen"></i>
          </button>
          <div id="eventModalBody"></div>
          <div class="modal-footer">
            <button id="editEventBtn"   class="btn btn-primary">Edit</button>
            <button id="deleteEventBtn" class="btn btn-danger">Delete</button>
          </div>
        </div>
      </div>
    `;

    // Toggle between minimal and full details
    modal
      .querySelector('.toggle-fullscreen')
      .addEventListener('click', () => {
        const dialog = modal.querySelector('.modal-dialog');
        const full   = dialog.classList.toggle('modal-fullscreen-custom');
        const data   = modal.dataset.eventData && JSON.parse(modal.dataset.eventData);

        if (!data) return;

        if (full) {
          const entry = window.entries.find(e => String(e.id) === modal.dataset.entryId);
          if (!entry) return;
          document.getElementById('eventModalBody').innerHTML = `
            <p><strong>Title:</strong>       ${entry.title}</p>
            <p><strong>Status:</strong>      ${getStatusBadge(entry.status)}</p>
            <p><strong>Start:</strong>       ${formatDate(entry.start)}</p>
            <p><strong>End:</strong>         ${formatDate(entry.end)}</p>
            <p><strong>Location:</strong>    ${entry.location}</p>
            <p><strong>Speaker:</strong>     ${entry.AssignedEmployee || '<em>Unassigned</em>'}</p>
            <p><strong>Topic:</strong>       ${entry.topic}</p>
            <p><strong>Industry:</strong>    ${entry.industry}</p>
            <p><strong>Type:</strong>        ${getTypePill(entry.internalExternal)}</p>
            <p><strong>Deadline:</strong>    ${formatDate(entry.applicationdeadline)}</p>
            <p><strong>Description:</strong> ${entry.desc}</p>
            <p><strong>Notes:</strong>       ${entry.notes}</p>
            <p><strong>Link:</strong>        ${
              entry.link
                ? `<a href="${entry.link}" target="_blank">${entry.link}</a>`
                : 'N/A'
            }</p>
          `;
        } else {
          document.getElementById('eventModalBody').innerHTML = `
            <p><strong>Title:</strong>    ${data.title}</p>
            <p><strong>Status:</strong>   ${getStatusBadge(data.status)}</p>
            <p><strong>Start:</strong>    ${formatDate(data.start)}</p>
            <p><strong>End:</strong>      ${formatDate(data.end)}</p>
            <p><strong>Location:</strong> ${data.location}</p>
            <p><strong>Speaker:</strong>  ${data.speaker}</p>
          `;
        }
      });

    return modal;
  }

  function getOrCreateEventModal() {
    let modal = document.getElementById('eventModal');
    if (!modal) {
      modal = createEventModal();
      document.body.appendChild(modal);
    }
    return modal;
  }
  function editEntryFromCalendar(entryId) {
    const modalEl = document.getElementById('eventModal');
    const bsModal = bootstrap.Modal.getInstance(modalEl);
    if (bsModal) bsModal.hide();

    document.getElementById('list-view-btn').click();

    setTimeout(() => {
      document.querySelectorAll('#entries-tbody tr').forEach(row => {
        const btn = row.querySelector('.edit-entry');
        if (btn && btn.dataset.id === String(entryId)) {
          btn.click();
        }
      });
    }, 200);
  }

  async function openEntryModalFromList(entry) {
    const modal = getOrCreateEventModal();

    modal.querySelector('#eventModalBody').innerHTML = `
      <p><strong>Title:</strong>       ${entry.title}</p>
      <p><strong>Topic:</strong>       ${entry.topic}</p>
      <p><strong>Status:</strong>      ${getStatusBadge(entry.status)}</p>
      <p><strong>Start:</strong>       ${formatDate(entry.start)}</p>
      <p><strong>End:</strong>         ${formatDate(entry.end)}</p>
      <p><strong>Location:</strong>    ${entry.location}</p>
      <p><strong>Speaker:</strong>     ${entry.AssignedEmployee || '<em>Unassigned</em>'}</p>
      <p><strong>Industry:</strong>    ${entry.industry}</p>
      <p><strong>Type:</strong>        ${getTypePill(entry.internalExternal)}</p>
      <p><strong>Deadline:</strong>    ${formatDate(entry.applicationdeadline)}</p>
      <p><strong>Description:</strong> ${entry.desc}</p>
      <p><strong>Notes:</strong>       ${entry.notes}</p>
      <p><strong>Link:</strong>        ${
        entry.link
          ? `<a href="${entry.link}" target="_blank">${entry.link}</a>`
          : 'N/A'
      }</p>
    `;

    modal.dataset.entryId = entry.id;
    modal.querySelector('#editEventBtn').onclick   = () => openEditModal(entry);
    modal.querySelector('#deleteEventBtn').onclick = async () => {
      if (!confirm(`Delete "${entry.title}"?`)) return;
      bootstrap.Modal.getInstance(modal).hide();
      showLoading();
      try {
        const token = await getToken(['Sites.ReadWrite.All']);
        const g     = client(token);
        const sid   = await getSiteId(g);
        await g
          .api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items/${entry.id}`)
          .delete();
        await fetchEntries();
      } catch (err) {
        console.error(err);
        alert('Error deleting entry: ' + err.message);
      } finally {
        hideLoading();
      }
    };

    new bootstrap.Modal(modal).show();
  }

  function renderList() {
  const tbody = document.getElementById('entries-tbody');
  const thead = document.getElementById('entries-thead');
  thead.innerHTML = '';

  // 1) Define which columns to show (and their object-keys)
  const columns = isCollapsedMode
    ? [
        { label: '#',      key: null,               width: '50' },
        { label: 'Title',  key: 'title'                },
        { label: 'Status', key: 'status'               },
        { label: 'Start',  key: 'start'                },
        { label: 'End',    key: 'end'                  },
        { label: 'Location', key: 'location'           },
        { label: 'Speaker',  key: 'AssignedEmployee'   }
      ]
    : [
        { label: '#',               key: null,               width: '50' },
        { label: 'Title',           key: 'title'                },
        { label: 'Topic',           key: 'topic'                },
        { label: 'Status',          key: 'status'               },
        { label: 'Start',           key: 'start'                },
        { label: 'End',             key: 'end'                  },
        { label: 'Location',        key: 'location'             },
        { label: 'Link',            key: 'link'                 },
        { label: 'Industry',        key: 'industry'             },
        { label: 'Description',     key: 'desc'                 },
        { label: 'Deadline',        key: 'applicationdeadline'  },
        { label: 'Internal/External', key: 'internalExternal'  },
        { label: 'Speaker',         key: 'AssignedEmployee'     },
        { label: 'Notes',           key: 'notes'                }
      ];

  // 2) Build the <thead> row with clickable <th>s
  const headRow = document.createElement('tr');
  columns.forEach(col => {
    const th = document.createElement('th');
    if (col.width) th.setAttribute('width', col.width);

    // always show the label
    th.appendChild(document.createTextNode(col.label + ' '));

    if (col.key) {
      // existing sort behavior for all other columns
      th.style.cursor = 'pointer';
      const icon = document.createElement('i');
      icon.classList.add('sort-icon', 'bi');
      if (sortColumn === col.key) {
        icon.classList.add(
          sortDirection === 'asc'
            ? 'bi-sort-alpha-down-alt'
            : 'bi-sort-alpha-up-alt'
        );
      } else {
        icon.classList.add('bi-sort-alpha-down');
      }
      th.appendChild(icon);
      th.addEventListener('click', () => {
        if (sortColumn === col.key) {
          sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
          sortColumn    = col.key;
          sortDirection = 'asc';
        }
        renderList();
      });
    }

    headRow.appendChild(th);
  });
  thead.appendChild(headRow);

  // 3) Sort a copy of entries if needed
  let entriesToShow = [...window.entries];
  if (sortColumn) {
    entriesToShow.sort((a, b) => {
      let va = a[sortColumn] ?? '';
      let vb = b[sortColumn] ?? '';

      // if date‐type columns, compare as Date objects
      if (['start','end','applicationdeadline'].includes(sortColumn)) {
        va = va ? new Date(va) : new Date(0);
        vb = vb ? new Date(vb) : new Date(0);
      } else {
        // string compare, case‐insensitive
        va = va.toString().toLowerCase();
        vb = vb.toString().toLowerCase();
      }

      if (va > vb) return sortDirection === 'asc' ?  1 : -1;
      if (va < vb) return sortDirection === 'asc' ? -1 :  1;
      return 0;
    });
  }

  // 4) Render rows (or empty state)
  tbody.innerHTML = '';
  if (!entriesToShow.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="${columns.length}" class="empty-state">
          <i class="bi bi-calendar-x"></i>
          <p>No conferences scheduled yet.</p>
        </td>
      </tr>
    `;
    return;
  }

  entriesToShow.forEach((e, idx) => {
    const tr = document.createElement('tr');
    if (isCollapsedMode) {
      tr.innerHTML = `
        <td>${idx+1}</td>
        <td><strong>${e.title}</strong></td>
        <td>${getStatusBadge(e.status)}</td>
        <td>${formatDate(e.start)}</td>
        <td>${formatDate(e.end)}</td>
        <td>${e.location}</td>
        <td>${e.AssignedEmployee || '<em>Unassigned</em>'}</td>
      `;
    } else {
      tr.innerHTML = `
        <td>${idx+1}</td>
        <td><strong>${e.title}</strong></td>
        <td>${e.topic}</td>
        <td>${getStatusBadge(e.status)}</td>
        <td>${formatDate(e.start)}</td>
        <td>${formatDate(e.end)}</td>
        <td>${e.location}</td>
        <td>${e.link ? `<a href="${e.link}" target="_blank"><i class="bi bi-link-45deg"></i></a>` : ''}</td>
        <td>${e.industry}</td>
        <td><small>${truncate(e.desc,50)}</small></td>
        <td>${formatDate(e.applicationdeadline)}</td>
        <td>${getTypePill(e.internalExternal)}</td>
        <td>${e.AssignedEmployee||'<em>Unassigned</em>'}</td>
        <td><small>${truncate(e.notes,50)}</small></td>
      `;
    }
    tr.addEventListener('dblclick', () => openEntryModalFromList(e));
    tbody.appendChild(tr);
  });
}


  function applyFilters() {
    const selectedSpeaker = document.getElementById('employee-filter').value;
    const month           = document.getElementById('start-date-filter').value;
    const query           = document.getElementById('keyword-search').value.trim().toLowerCase();
    const typeFilter      = document.getElementById('internal-external-filter').value;
    const rows            = document.querySelectorAll('#entries-tbody tr');

    rows.forEach((row, i) => {
      const entry = window.entries[i];
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
// new internal/external filter
      if (typeFilter && entry.internalExternal !== typeFilter) show = false;
      row.style.display = show ? '' : 'none';
    });
  }

  // ---- EXISTING BUTTON & VIEW HANDLERS ----
  document.getElementById('edit-employees-btn')
    .addEventListener('click', () =>
      new bootstrap.Modal(document.getElementById('employeeModal')).show()
    );
  document.getElementById('add-emp-btn').addEventListener('click', addEmployee);
  document.getElementById('open-add-entry-ui-btn').addEventListener('click', () => {
    const modalEl    = document.getElementById('addEntryModal');
    const addModal   = new bootstrap.Modal(modalEl);
    addModal.show();
  });
  document.getElementById('cancel-entry-btn').addEventListener('click', () => {
    clearAddForm();
    const modalEl = document.getElementById('addEntryModal');
    bootstrap.Modal.getInstance(modalEl).hide();
  });
  document.getElementById('submit-entry-btn').addEventListener('click', async () => {
    await addEntry();
    clearAddForm();
    const modalEl = document.getElementById('addEntryModal');
    bootstrap.Modal.getInstance(modalEl).hide();
  });
document.getElementById('edit-submit-btn').addEventListener('click', updateEntry);
  document.getElementById('employee-filter').addEventListener('change', applyFilters);
  document.getElementById('start-date-filter').addEventListener('change', applyFilters);
  document.getElementById('keyword-search').addEventListener('input', applyFilters);
  document.getElementById('internal-external-filter').addEventListener('change', applyFilters);
  document.getElementById('calendar-employee-filter').addEventListener('change', renderCalendarView);

  calBtn.addEventListener('click', () => {
    listBtn.classList.remove('active');
    calBtn.classList.add('active');
    listContainer.classList.add('hidden');
    calendarContainer.classList.remove('hidden');
    renderCalendarView();
  });

  listBtn.addEventListener('click', () => {
    calBtn.classList.remove('active');
    listBtn.classList.add('active');
    calendarContainer.classList.add('hidden');
    listContainer.classList.remove('hidden');
    renderList();
  });

  document.getElementById('calendar-list-view-btn')
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
