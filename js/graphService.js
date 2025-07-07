// js/graphService.js
// …above your initLists/fetchEmployees/fetchEntries…

/** Exported so events.js can import it */
export async function addEmployee() {
  const name  = prompt('Enter employee name:');
  const email = prompt('Enter employee email:');
  if (!name || !email) return;
  showLoading();
  const token = await getToken(['Sites.ReadWrite.All']);
  const g     = client(token);
  const sid   = await getSiteId(g);
  await g.api(`/sites/${sid}/lists/${EMP_LIST_ID}/items`)
         .post({ fields: { Title: name, Employeesemail: email } });
  await fetchEmployees();
  hideLoading();
}

/** Exported so events.js can import it */
export async function addEntry() {
  showLoading();
  try {
    const token = await getToken(['Sites.ReadWrite.All']);
    const g     = client(token);
    const sid   = await getSiteId(g);

    // gather your form values…
    const fields = {
      Title:                 document.getElementById('new-title').value,
      Topic:                 document.getElementById('new-topic').value,
      RegistrationStatus:    document.getElementById('new-status').value,
      StartDate:             document.getElementById('new-start').value,
      EndDate:               document.getElementById('new-end').value,
      Location:              document.getElementById('new-location').value,
      Link:                  document.getElementById('new-link').value,
      Industry:              document.getElementById('new-industry').value,
      Description:           document.getElementById('new-desc').value,
      ApplicationDeadline:   document.getElementById('new-applicationdeadline').value,
      Internal_x002f_External: document.getElementById('new-type').value,
      AssignedEmployee:        document.getElementById('new-AssignedEmployee').value,
      Notes:                   document.getElementById('new-notes').value
    };

    await g.api(`/sites/${sid}/lists/${SCHED_LIST_ID}/items`)
           .post({ fields });

    // optional: calendar invite logic…
    await fetchEntries();
    clearAddForm();    // if you have one
  } catch (err) {
    console.error(err);
    alert(err.message);
  } finally {
    hideLoading();
  }
}

