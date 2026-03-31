// ── Tab navigation ──
document.querySelectorAll('.tab').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
    if (btn.dataset.tab === 'tickets') loadTickets();
  });
});

// ── Redline ──
const redlineForm = document.getElementById('redline-form');
const redlineStatus = document.getElementById('redline-status');

function showStatus(el, msg, type) {
  el.textContent = msg;
  el.className = `status ${type}`;
  el.hidden = false;
}

redlineForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const formData = new FormData(redlineForm);
  const btn = document.getElementById('btn-compare');
  btn.disabled = true;
  showStatus(redlineStatus, 'Comparing documents...', 'info');

  try {
    const resp = await fetch('/api/compare', { method: 'POST', body: formData });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ detail: resp.statusText }));
      throw new Error(err.detail || err.error || 'Comparison failed');
    }
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'redline.docx';
    a.click();
    URL.revokeObjectURL(url);
    showStatus(redlineStatus, 'Redline downloaded!', 'success');
  } catch (err) {
    showStatus(redlineStatus, err.message, 'error');
  } finally {
    btn.disabled = false;
  }
});

document.getElementById('btn-preview').addEventListener('click', async () => {
  const formData = new FormData(redlineForm);
  const btn = document.getElementById('btn-preview');
  btn.disabled = true;
  showStatus(redlineStatus, 'Generating HTML preview...', 'info');
  const preview = document.getElementById('html-preview');

  try {
    const resp = await fetch('/api/compare/html', { method: 'POST', body: formData });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ detail: resp.statusText }));
      throw new Error(err.detail || err.error || 'Comparison failed');
    }
    const html = await resp.text();
    preview.innerHTML = '';
    const iframe = document.createElement('iframe');
    preview.appendChild(iframe);
    preview.hidden = false;
    iframe.srcdoc = html;
    showStatus(redlineStatus, 'Preview ready.', 'success');
  } catch (err) {
    showStatus(redlineStatus, err.message, 'error');
    preview.hidden = true;
  } finally {
    btn.disabled = false;
  }
});

// ── Tickets list ──
let currentPage = 1;

async function loadTickets(page = 1) {
  currentPage = page;
  const status = document.getElementById('ticket-filter').value;
  const params = new URLSearchParams({ page, pageSize: 25 });
  if (status) params.set('status', status);

  try {
    const resp = await fetch(`/api/tickets?${params}`);
    const data = await resp.json();
    const tbody = document.querySelector('#ticket-table tbody');
    tbody.innerHTML = '';

    for (const t of data.tickets) {
      const tr = document.createElement('tr');
      tr.addEventListener('click', () => openTicketModal(t.id));
      const badgeClass = 'badge-' + t.status.toLowerCase().replace(/\s+/g, '');
      tr.innerHTML = `
        <td>${t.id}</td>
        <td>${esc(t.title)}</td>
        <td><span class="badge ${badgeClass}">${t.status}</span></td>
        <td>${t.revisionCount ?? '—'}</td>
        <td>${new Date(t.createdAt).toLocaleDateString()}</td>
        <td>
          <a href="/api/tickets/${t.id}/files/original" onclick="event.stopPropagation()">Original</a> ·
          <a href="/api/tickets/${t.id}/files/modified" onclick="event.stopPropagation()">Modified</a>
          ${t.revisionCount != null ? ` · <a href="/api/tickets/${t.id}/files/redline" onclick="event.stopPropagation()">Redline</a>` : ''}
        </td>
      `;
      tbody.appendChild(tr);
    }

    // Paging
    const paging = document.getElementById('ticket-paging');
    const totalPages = Math.ceil(data.total / data.pageSize);
    paging.innerHTML = '';
    for (let p = 1; p <= totalPages; p++) {
      const btn = document.createElement('button');
      btn.textContent = p;
      btn.disabled = p === currentPage;
      btn.addEventListener('click', () => loadTickets(p));
      paging.appendChild(btn);
    }
  } catch (err) {
    console.error('Failed to load tickets', err);
  }
}

document.getElementById('ticket-filter').addEventListener('change', () => loadTickets(1));
document.getElementById('btn-refresh').addEventListener('click', () => loadTickets(currentPage));

// ── Ticket detail modal ──
const modal = document.getElementById('ticket-modal');
let currentTicketId = null;

async function openTicketModal(id) {
  currentTicketId = id;
  const resp = await fetch(`/api/tickets/${id}`);
  const t = await resp.json();

  document.getElementById('modal-title').textContent = `#${t.id} — ${t.title}`;
  document.getElementById('modal-status').value = t.status;

  let html = `
    <p><strong>Description:</strong> ${esc(t.description) || '<em>None</em>'}</p>
    <p><strong>Submitter:</strong> ${esc(t.submitterEmail) || '<em>Anonymous</em>'}</p>
    <p><strong>Created:</strong> ${new Date(t.createdAt).toLocaleString()}</p>
    <p><strong>Updated:</strong> ${new Date(t.updatedAt).toLocaleString()}</p>
    <p><strong>Revisions detected:</strong> ${t.revisionCount ?? 'N/A'}</p>
    <p><strong>Files:</strong>
      <a href="/api/tickets/${t.id}/files/original">${esc(t.originalFileName)}</a> ·
      <a href="/api/tickets/${t.id}/files/modified">${esc(t.modifiedFileName)}</a>
      ${t.revisionCount != null ? ` · <a href="/api/tickets/${t.id}/files/redline">Redline</a>` : ''}
    </p>
  `;
  if (t.comparisonLog) {
    html += `<p><strong>Comparison Log:</strong></p><div class="log">${esc(t.comparisonLog)}</div>`;
  }
  document.getElementById('modal-body').innerHTML = html;
  modal.showModal();
}

document.getElementById('btn-close-modal').addEventListener('click', () => modal.close());
modal.addEventListener('click', (e) => { if (e.target === modal) modal.close(); });

document.getElementById('btn-update-status').addEventListener('click', async () => {
  const status = document.getElementById('modal-status').value;
  await fetch(`/api/tickets/${currentTicketId}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ status }),
  });
  modal.close();
  loadTickets(currentPage);
});

// ── Submit ticket ──
const ticketForm = document.getElementById('ticket-form');
const submitStatus = document.getElementById('submit-status');

ticketForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const formData = new FormData(ticketForm);
  const btn = document.getElementById('btn-submit-ticket');
  btn.disabled = true;
  showStatus(submitStatus, 'Uploading files and running comparison...', 'info');

  try {
    const resp = await fetch('/api/tickets', { method: 'POST', body: formData });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ detail: resp.statusText }));
      throw new Error(err.detail || err.error || 'Submission failed');
    }
    const data = await resp.json();
    let msg = `Ticket #${data.id} created.`;
    if (data.revisionCount != null) msg += ` ${data.revisionCount} revision(s) detected.`;
    if (data.comparisonLog) msg += ` (see ticket for comparison warnings)`;
    showStatus(submitStatus, msg, 'success');
    ticketForm.reset();
  } catch (err) {
    showStatus(submitStatus, err.message, 'error');
  } finally {
    btn.disabled = false;
  }
});

// ── Util ──
function esc(s) {
  if (!s) return '';
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}
