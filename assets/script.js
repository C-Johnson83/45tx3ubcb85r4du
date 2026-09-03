// EDIT THIS ARRAY to add your real scenarios.
    // Search matches anywhere inside the id, title, summary, description, solution, or keywords.
    const scenarios = [
      {
        id: "S-001",
        title: "Bulk Reinstating 1",
        summary: "Location is on Hold Both. It is systematically empty, but there is physical freight!",
        description: "The location will become available for IB as soon as the hold is released. Procaution must be used to prevent IB from using it for a receipt as soon as we release it.",
        solution: "If the location is on hold, and the location shows empty in WLI, to reinstate the product back in the location, it must be done quickly. Scan the product in ISI/IUP and jump to PAR. Fill out the screen including the Aisle and Bin but rest on the Level. On the equipments screen pull the location up in WLI and press F4. Be ready with the handheld. Enter RH on the equipment screen and press F1. Quickly, on the Telxon screen, enter 01 and press F1",
        keywords: ["Bulk", "Reinstate", "Max Pallet"]
    }

,
    {
        id: "S-002",
        title: "Bulk Reinstating 2",
        summary: "Location is on Hold Both. It is systematically Occupied, but there is no physical freight!",
        description: "The location has most likely already been used for a receipt. The carton count will not show until a PUT event has been completed.",
        solution: "If the location is on hold, and the location shows empty in WLI, to reinstate the product back in the location, it must be done quickly. Scan the product in ISI/IUP and jump to PAR. Fill out the screen including the Aisle and Bin but rest on the Level. On the equipments screen pull the location up in WLI and press F4. Be ready with the handheld. Enter RH on the equipment screen and press F1. Quickly, on the Telxon screen, enter 01 and press F1",
        keywords: ["Bulk", "Reinstate", "Max Pallet"]
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
