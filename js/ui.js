// js/ui.js
import * as utils from './utils.js';

export function populateEmployeeSelects() {
  const sel    = document.getElementById('new-AssignedEmployee');
  const filter = document.getElementById('employee-filter');
  const cal    = document.getElementById('calendar-employee-filter');
  [sel,filter,cal].forEach(s=> s.innerHTML='<option value="">All Employees</option>');

  window.employees.forEach(e=>{
    [sel,filter,cal].forEach(s=>{
      const o = document.createElement('option');
      o.value = e.name; o.textContent = e.name;
      s.appendChild(o);
    });
  });
}

export function renderEmployeeTable() {
  // …copy your original renderEmployeeTable() here…
}

export function renderList() {
  // …copy your original renderList() + applyFilters() + enableEditMode() here…
}

export function renderCalendarView() {
  // …copy your original FullCalendar-init logic here…
}
