let services = [];
let fuse = null;
let currentCategory = 'all';
let bookmarks = JSON.parse(localStorage.getItem('bookmarks') || '[]');

async function loadServices() {
    try {
        const res = await fetch('services.json');
        services = await res.json();
        fuse = new Fuse(services, {
            keys: ['name', 'category', 'description', 'tags'],
            threshold: 0.3,
            includeScore: true
        });
        renderServices(services);
    } catch (err) {
        console.error('Failed to load services:', err);
        document.getElementById('results').innerHTML = '<p class="no-results">데이터를 불러오는데 실패했습니다.</p>';
    }
}

function renderServices(results) {
    const container = document.getElementById('results');
    const countEl = document.getElementById('resultCount');
    
    if (!results.length) {
        container.innerHTML = '<p class="no-results">검색 결과가 없습니다.</p>';
        countEl.textContent = '0';
        return;
    }
    
    countEl.textContent = results.length;
    container.innerHTML = results.map((svc, i) => `
        <div class="service-card">
            <div class="service-header">
                <div>
                    <span class="service-name">
                        <a href="${svc.url}" target="_blank" rel="noopener">${svc.name}</a>
                    </span>
                    <span class="service-category">${svc.category}</span>
                </div>
                <button class="bookmark-btn ${bookmarks.includes(svc.name) ? 'active' : ''}" 
                        onclick="toggleBookmark('${svc.name}')" title="북마크">★</button>
            </div>
            <p class="service-description">${svc.description}</p>
            <div class="service-tags">
                ${svc.tags.map(t => `<span class="tag">#${t}</span>`).join('')}
            </div>
        </div>
    `).join('');
}

function toggleBookmark(name) {
    const idx = bookmarks.indexOf(name);
    if (idx > -1) {
        bookmarks.splice(idx, 1);
    } else {
        bookmarks.push(name);
    }
    localStorage.setItem('bookmarks', JSON.stringify(bookmarks));
    const currentResults = getFilteredResults();
    renderServices(currentResults);
}

function getFilteredResults() {
    const query = document.getElementById('searchInput').value.trim();
    
    if (!query) {
        if (currentCategory === 'all') return services;
        return services.filter(s => s.category === currentCategory);
    }
    
    let results = fuse.search(query).map(r => r.item);
    if (currentCategory !== 'all') {
        results = results.filter(s => s.category === currentCategory);
    }
    return results;
}

document.getElementById('searchInput').addEventListener('input', () => {
    renderServices(getFilteredResults());
});

document.querySelectorAll('.category-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.category-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentCategory = btn.dataset.category;
        renderServices(getFilteredResults());
    });
});

// Init
loadServices();