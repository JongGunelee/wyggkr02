const CACHE_NAME = 'automation-dashboard-v1';
const CORE_ASSETS = [
    './',
    './index.html',
    './00%20dashboard.html',
    './manifest.webmanifest',
    './000%20Launch_dashboard.bat',
    './run_dashboard.py',
    './automated_scripts/group_cross_merger.py',
    './automated_scripts/pdf_to_html_converter_ultimate.py',
    './automated_scripts/Batch_PPT_to_PDF_DDD.py',
    './automated_scripts/advanced_pdf_compressor.py',
    './automated_scripts/universal_office_optimizer.py',
    './automated_scripts/excel_deep_cleaner.py',
    './automated_scripts/modify_excel_repair.py',
    './automated_scripts/advanced_excel_rename.py',
    './automated_scripts/advanced_column_modifier.py',
    './automated_scripts/search_two_items.py',
    './automated_scripts/batch_copy_pdf.py',
    './automated_scripts/pattern_document_merger.py',
    './automated_scripts/intelligent_file_organizer.py',
    './automated_scripts/excel_compressor_tool.py',
    './automated_scripts/ppt_compressor_tool.py',
    './automated_scripts/collect_closing_data.py',
    './system_guides/00%20PRD%20가이드.md',
    './system_guides/AI_CODING_GUIDELINES_2026.md',
    './system_guides/Office_Stability_Audit_Report.md',
    './system_guides/000%20프롬프트(스크립트).txt'
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS)).then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((keys) =>
            Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))
        ).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', (event) => {
    if (event.request.method !== 'GET') {
        return;
    }

    if (event.request.url.startsWith('http://127.0.0.1:8501')) {
        return;
    }

    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            if (cachedResponse) {
                return cachedResponse;
            }

            return fetch(event.request).then((networkResponse) => {
                if (!networkResponse || !networkResponse.ok || !event.request.url.startsWith(self.location.origin)) {
                    return networkResponse;
                }

                const responseClone = networkResponse.clone();
                caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
                return networkResponse;
            }).catch(() => caches.match('./00%20dashboard.html'));
        })
    );
});
