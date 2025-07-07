// js/utils.js
export function showLoading() {
  document.getElementById('loading-bar').classList.add('active');
}
export function hideLoading() {
  setTimeout(() => {
    document.getElementById('loading-bar').classList.remove('active');
  }, 300);
}

export function formatDate(iso) {
  if (!iso) return '';
  return new Date(iso).toLocaleDateString('en-US',{
    month:'short', day:'numeric', year:'numeric'
  });
}
export function formatDateOnly(iso) {
  return iso ? iso.split('T')[0] : '';
}
export function truncate(str,len) {
  return str && str.length>len
    ? str.slice(0,len)+'…'
    : (str||'');
}

export function getStatusBadge(status) {
  const cls = {
    Confirmed:'status-confirmed',
    Pending:  'status-pending',
    Cancelled:'status-cancelled',
    Completed:'status-completed'
  }[status]||'status-pending';
  return `<span class="status-badge ${cls}">${status||'Pending'}</span>`;
}

export function getTypePill(type) {
  const cls = type==='Internal' ? 'type-internal' : 'type-external';
  return `<span class="type-pill ${cls}">${type||'Internal'}</span>`;
}
