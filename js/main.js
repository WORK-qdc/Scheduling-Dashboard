// js/main.js
import { initLists, fetchEmployees, fetchEntries } from './graphService.js';
import * as ui    from './ui.js';
import { bindUI } from './events.js';

document.addEventListener('DOMContentLoaded', async () => {
  try {
    await initLists();
    await fetchEmployees();
    ui.populateEmployeeSelects();
    ui.renderEmployeeTable();

    await fetchEntries();
    ui.renderList();

    bindUI();
  } catch (err) {
    console.error('Failed to start app:', err);
    alert('Initialization error:\n' + err.message);
  }
});
