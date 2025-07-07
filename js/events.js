// js/events.js

import { addEmployee, addEntry } from './graphService.js';
import * as ui from './ui.js';

export function bindUI() {
  // Open Employee Management modal
  document
    .getElementById('edit-employees-btn')
    .addEventListener('click', () => {
      new bootstrap.Modal(document.getElementById('employeeModal')).show();
    });

  // Add a new employee
  document
    .getElementById('add-emp-btn')
    .addEventListener('click', addEmployee);

  // Add a new conference entry
  document
    .getElementById('add-entry-btn')
    .addEventListener('click', addEntry);

  // Switch to List View
  document
    .getElementById('list-view-btn')
    .addEventListener('click', () => {
      document.getElementById('calendar-view-btn').classList.remove('active');
      document.getElementById('list-view-btn').classList.add('active');
      document.getElementById('calendar-container').style.display = 'none';
      document.getElementById('list-container').style.display     = 'block';
    });

  // Switch to Calendar View
  document
    .getElementById('calendar-view-btn')
    .addEventListener('click', () => {
      document.getElementById('list-view-btn').classList.remove('active');
      document.getElementById('calendar-view-btn').classList.add('active');
      document.getElementById('list-container').style.display     = 'none';
      document.getElementById('calendar-container').style.display = 'block';
      ui.renderCalendarView();
    });

  // Back to List from Calendar
  document
    .getElementById('calendar-list-view-btn')
    .addEventListener('click', () => {
      document.getElementById('list-view-btn').click();
    });

  // Filters in List View
  document
    .getElementById('employee-filter')
    .addEventListener('change', ui.applyFilters);

  document
    .getElementById('start-date-filter')
    .addEventListener('change', ui.applyFilters);

  // Filter in Calendar View
  document
    .getElementById('calendar-employee-filter')
    .addEventListener('change', ui.renderCalendarView);
}
