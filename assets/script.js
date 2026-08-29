// EDIT THIS ARRAY to add your real scenarios.
    // Search matches anywhere inside the id, title, summary, description, solution, or keywords.
    const scenarios = [
      {
        id: "S-001",
        title: "Bulk Reinstating",
        summary: "Location is on Hold Both. It is systematically empty, but there is physical freight!",
        description: "The location will become available for IB as soon as the hold is released. Procaution must be used to prevent IB from using it for a receipt as soon as we release it.",
        solution: "If the location is on hold, and the location shows empty in WLI, to reinstate the product back in the location, it must be done quickly. Scan the product in ISI/IUP and jump to PAR. Fill out the screen including the Aisle and Bin but rest on the Level. On the equipments screen pull the location up in WLI and press F4. Be ready with the handheld. Enter RH on the equipment screen and press F1. Quickly, on the Telxon screen, enter 01 and press F1",
        keywords: ["Bulk", "Reinstate", "Max Pallet"]
      },
      {
        id: "S-002",
        title: "Order shows delivered but is missing",
        summary: "Tracking indicates delivery, but the customer cannot locate the package.",
        description: "The shipment has a delivered scan, yet the customer says the item is not at the expected location. Check the scan time, delivery location details, household members, and any carrier notes.",
        solution: "Confirm the delivery details with the customer, ask them to check nearby secure locations, and follow the approved missing-delivery workflow. If the required waiting period has passed, proceed with the authorized replacement, refund, or carrier claim process.",
        keywords: ["delivered", "missing", "package", "tracking", "carrier", "order"]
      },
      {
        id: "S-003",
        title: "Payment was declined",
        summary: "A transaction is rejected during checkout or account payment.",
        description: "The customer receives a decline message when submitting payment. The cause may involve billing information, card restrictions, available funds, processor rules, or a temporary authorization issue.",
        solution: "Confirm billing details without collecting prohibited sensitive data, retry only when policy allows, and offer an approved alternate payment method. For persistent processor errors, capture the non-sensitive error code and follow the payment escalation procedure.",
        keywords: ["payment", "declined", "card", "billing", "checkout", "transaction"]
      },
      {
        id: "S-004",
        title: "Customer requests a supervisor",
        summary: "The customer asks to speak with leadership or requests escalation.",
        description: "The customer may be dissatisfied with the proposed resolution, may repeat a request for a supervisor, or may explicitly ask for escalation. Avoid debating the request and document the reason clearly.",
        solution: "Acknowledge the request, complete any required verification, summarize what has already been attempted, and use the approved escalation path. Set accurate expectations about response or transfer timing.",
        keywords: ["supervisor", "manager", "escalate", "escalation", "leadership", "complaint"]
      },
      {
        id: "S-005",
        title: "Website page will not load",
        summary: "A user reports a blank page, repeated loading, or a page error.",
        description: "The issue may be isolated to one browser, one device, or a particular page. Determine whether the user sees an error message and whether other pages are working normally.",
        solution: "Confirm the affected URL or workflow, have the user refresh or retry using the approved troubleshooting steps, and check for a known outage. If reproducible, document browser/device details and escalate with the error text or screenshot when permitted.",
        keywords: ["website", "page", "blank", "loading", "browser", "error", "outage"]
      }
    ];

    const app = document.getElementById('app');
    const list = document.getElementById('scenarioList');
    const search = document.getElementById('search');
    const count = document.getElementById('count');
    const collapseBtn = document.getElementById('collapseBtn');
    let selectedId = scenarios[0]?.id;

    function haystack(s){
      return [s.id,s.title,s.summary,s.description,s.solution,...s.keywords].join(' ').toLowerCase();
    }

    function renderList(query=''){
      const q = query.trim().toLowerCase();
      const filtered = q ? scenarios.filter(s => haystack(s).includes(q)) : scenarios;
      count.textContent = `${filtered.length} of ${scenarios.length} scenarios`;
      list.innerHTML = '';

      if(!filtered.length){
        list.innerHTML = '<div class="empty">No matching scenarios</div>';
        return;
      }

      filtered.forEach(s => {
        const btn = document.createElement('button');
        btn.className = 'scenario-item' + (s.id === selectedId ? ' active' : '');
        btn.innerHTML = `<strong>${escapeHtml(s.title)}</strong><span>${escapeHtml(s.keywords.join(' · '))}</span>`;
        btn.onclick = () => { selectedId = s.id; showScenario(s); renderList(search.value); };
        list.appendChild(btn);
      });

      if(q && !filtered.some(s => s.id === selectedId)){
        selectedId = filtered[0].id;
        showScenario(filtered[0]);
        renderList(search.value);
      }
    }

    function showScenario(s){
      document.getElementById('scenarioId').textContent = s.id;
      document.getElementById('scenarioTitle').textContent = s.title;
      document.getElementById('scenarioSummary').textContent = s.summary;
      document.getElementById('description').textContent = s.description;
      document.getElementById('solution').textContent = s.solution;
      document.getElementById('keywords').innerHTML = s.keywords.map(k => `<span class="chip">${escapeHtml(k)}</span>`).join('');
    }

    function escapeHtml(value){
      return String(value).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
    }

    search.addEventListener('input', e => renderList(e.target.value));
    search.addEventListener('keydown', e => {
      if(e.key === 'Escape'){ search.value=''; renderList(); search.blur(); }
      if(e.key === 'Enter'){
        const q = search.value.trim().toLowerCase();
        const first = scenarios.find(s => !q || haystack(s).includes(q));
        if(first){ selectedId=first.id; showScenario(first); renderList(search.value); }
      }
    });

    collapseBtn.addEventListener('click', () => {
      app.classList.toggle('sidebar-collapsed');
      collapseBtn.textContent = app.classList.contains('sidebar-collapsed') ? '▶' : '◀';
    });

    showScenario(scenarios[0]);
    renderList();
