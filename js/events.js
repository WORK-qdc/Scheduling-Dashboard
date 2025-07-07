// js/events.js
import { initLists, fetchEmployees, fetchEntries, addEmployee, addEntry } from './graphService.js';
import * as ui    from './ui.js';

export function bindUI() {
  document.getElementById('add-emp-btn')
    .addEventListener('click', addEmployee);

  document.getElementById('add-entry-btn')
    .addEventListener('click', addEntry);

  document.getElementById('list-view-btn')
    .addEventListener('click', () => {
      document.getElementById('calendar-container').style.display = 'none';
      document.getElementById('list-container').style.display     = 'block';
    });

  document.getElementById('calendar-view-btn')
    .addEventListener('click', () => {
      document.getElementById('list-container').style.display     = 'none';
      document.getElementById('calendar-container').style.display = 'block';
      ui.renderCalendarView();
    });

  document.getElementById('calendar-list-view-btn')
    .addEventListener('click', ()=> document.getElementById('list-view-btn').click());

  document.getElementById('employee-filter')
    .addEventListener('change', ui.applyFilters);

  document.getElementById('start-date-filter')
    .addEventListener('change', ui.applyFilters);

  document.getElementById('calendar-employee-filter')
    .addEventListener('change', ui.renderCalendarView);

  document.getElementById('edit-employees-btn')
    .addEventListener('click', ()=> new bootstrap.Modal(document.getElementById('employeeModal')).show());
}
