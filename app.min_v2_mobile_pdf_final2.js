const firebaseConfig = {
          apiKey: "AIzaSyDBdfswGWEefQR8Djscm2hkaKsi0438ODQ",
          authDomain: "webshop-japan.firebaseapp.com",
          projectId: "webshop-japan",
          storageBucket: "webshop-japan.firebasestorage.app",
          messagingSenderId: "854451712478",
          appId: "1:854451712478:web:d55198786d10a9b98d2d4f"
        };
        const appId = 'xnxl-offshore'; 
        if (!firebase.apps.length) firebase.initializeApp(firebaseConfig);
        const auth = firebase.auth();
        const db = firebase.firestore();
        
        const state = {
            currentUser: null,
            userRole: null, 
            userOrgId: null,
            currentDateKey: new Date().toISOString().split('T')[0].replace(/-/g, ''),
            orgs: [],
            isPlanLocked: false,
            currentPlanPeople: [],
            unsubscribePlan: null,
            users: [],
            quickSearch: ''
        };

        const utils = {
            formatDate: (dateStr) => { if(!dateStr || dateStr.length !== 8) return dateStr || ''; return `${dateStr.substr(6,2)}/${dateStr.substr(4,2)}/${dateStr.substr(0,4)}`; },
            formatDateTime: (dateObj) => { if(!dateObj) return ''; const d = new Date(dateObj); return `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')} ${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}`; },
            formatTimeShort: (dateObj) => { if(!dateObj) return ''; const d = new Date(dateObj); return `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`; },
            processExcelDate: (input) => {
                if (!input) return '';
                if (input instanceof Date) {
                    if (!isNaN(input.getTime())) {
                        return `${String(input.getDate()).padStart(2,'0')}/${String(input.getMonth()+1).padStart(2,'0')}/${input.getFullYear()}`;
                    }
                }
                if (typeof input === 'number') {
                    const date = new Date(Math.round((input - 25569) * 86400 * 1000));
                    if (!isNaN(date.getTime())) return `${String(date.getDate()).padStart(2,'0')}/${String(date.getMonth()+1).padStart(2,'0')}/${date.getFullYear()}`;
                }
                const str = String(input).trim();
                const parts = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
                if (parts) return `${parts[1].padStart(2,'0')}/${parts[2].padStart(2,'0')}/${parts[3]}`;
                return str;
            },
            showToast: (msg, type = 'info') => {
                const el = document.getElementById('toast');
                const colors = type === 'error' ? 'bg-red-500' : 'bg-emerald-600';
                const icon = type === 'error' ? '<i class="fas fa-exclamation-circle"></i>' : '<i class="fas fa-check-circle"></i>';
                el.className = `fixed top-6 right-6 z-[70] px-6 py-4 rounded-lg shadow-xl text-white font-medium transform transition-all duration-300 translate-y-0 opacity-100 flex items-center gap-3 ${colors}`;
                el.innerHTML = `${icon} <span>${msg}</span>`;
                el.classList.remove('hidden');
                setTimeout(() => { el.classList.add('opacity-0', '-translate-y-2'); setTimeout(() => el.classList.add('hidden', 'opacity-100', 'translate-y-0'), 300); }, 3000);
            },
            toggleLoader: (show, text) => {
                const el = document.getElementById('loadingOverlay');
                if(text) document.getElementById('loadingText').innerText = text;
                show ? el.classList.remove('hidden') : el.classList.add('hidden');
            },
            cleanString: (str) => (str || '').toString().trim(),

            // --- NVSX / Task aggregation (dedupe & smart merge) ---
            splitNvsxTokens: (val) => {
                if (!val) return [];
                return String(val)
                    .replace(/\r/g, '')
                    .split(/[\n;,]+/g)
                    .map(s => s.trim())
                    .filter(Boolean);
            },
            aggregateGroupNvsxAndTasks: (group) => {
                const nvsxOrder = [];
                const seenN = new Set();
                const taskMap = new Map(); // code -> Set(desc)
                const misc = new Set();

                (group || []).forEach(p => {
                    // 1) NVSX tokens from field
                    utils.splitNvsxTokens(p?.nvsxNo).forEach(code => {
                        const c = code.trim();
                        if (!c || seenN.has(c)) return;
                        seenN.add(c);
                        nvsxOrder.push(c);
                    });

                    // 2) Task lines
                    const raw = (p?.taskDesc || '').toString().replace(/\r/g, '');
                    if (!raw) return;
                    raw.split('\n').forEach(line => {
                        const l = (line || '').trim();
                        if (!l) return;

                        // Try parse "CODE: description" or "CODE - description"
                        let m = l.match(/^(\d{3}-\d{2}(?:bs\d+)?)\s*:\s*(.+)$/i);
                        if (!m) m = l.match(/^([0-9A-Za-z][0-9A-Za-z\-\/.]*)\s*:\s*(.+)$/);
                        if (!m) m = l.match(/^([0-9A-Za-z][0-9A-Za-z\-\/.]*)\s*-\s*(.+)$/);
                        if (m) {
                            const code = (m[1] || '').trim();
                            const desc = (m[2] || '').trim();
                            if (!code || !desc) return;

                            if (!taskMap.has(code)) taskMap.set(code, new Set());
                            taskMap.get(code).add(desc);
                        } else {
                            misc.add(l);
                        }
                    });
                });

                // Merge codes: NVSX order first, then codes only from taskMap
                const codesOnlyInTask = [...taskMap.keys()].filter(c => !seenN.has(c)).sort();
                const mergedCodes = [...nvsxOrder, ...codesOnlyInTask];

                // Build NVSX string
                const nvsxStr = mergedCodes.filter(Boolean).join('; ');

                // Build task lines (dedupe)
                const taskLines = [];
                mergedCodes.forEach(code => {
                    if (!taskMap.has(code)) return;
                    const descs = [...taskMap.get(code)];
                    if (descs.length === 1) taskLines.push(`${code}: ${descs[0]}`);
                    else taskLines.push(`${code}: ${descs.join(' / ')}`);
                });

                // Add misc lines (dedupe)
                [...misc].forEach(l => taskLines.push(l));

                return { nvsxStr: nvsxStr || '', taskLines };
            }
            ,
            normalizeNvsxInput: (val) => {
                // Accept separators: ';' ',' newline
                const tokens = utils.splitNvsxTokens(val);
                // Preserve original order but dedupe
                const seen = new Set();
                const out = [];
                tokens.forEach(t => { if (!seen.has(t)) { seen.add(t); out.push(t); } });
                return out.join('; ');
            },
            normalizeTaskText: (val) => {
                if (val == null) return '';
                return String(val).replace(/\r/g, '').trim();
            },
            normalizeDestinationKey: (dest) => {
                const raw = (dest ?? '').toString().trim();
                if (!raw) return '';
                const n = raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
                if (n === 'chua xac dinh') return '';
                return raw;
            },
            destinationDisplay: (dest) => {
                const key = utils.normalizeDestinationKey(dest);
                return key || 'Chưa xác định';
            },
            getGroupFinalMeta: (group) => {
                // Single source of truth = latest updated record (usually after Admin "Sửa giàn/NVSX/CV").
                // If no updatedAt, fall back to createdAt/importedAt ordering. If still missing, fall back to union aggregation.
                const pickTime = (p) => {
                    const t = p?.updatedAt || p?.createdAt || p?.importedAt || p?.importTime || null;
                    try {
                        if (!t) return 0;
                        if (t instanceof Date) return t.getTime();
                        if (typeof t === 'number') return t;
                        if (typeof t === 'string') return Date.parse(t) || 0;
                        if (t?.toDate) return t.toDate().getTime();
                        if (t?.seconds) return (t.seconds * 1000) + Math.floor((t.nanoseconds||0)/1e6);
                    } catch(e) {}
                    return 0;
                };

                let latest = null;
                let latestTs = -1;
                (group || []).forEach(p => {
                    const ts = pickTime(p);
                    if (ts > latestTs) { latestTs = ts; latest = p; }
                });

                let nvsxStr = utils.normalizeNvsxInput(latest?.nvsxNo || '');
                let taskText = utils.normalizeTaskText(latest?.taskDesc || '');

                // If latest is empty, fall back to union aggregation so UI still shows something.
                if (!nvsxStr && !taskText) {
                    const agg = utils.aggregateGroupNvsxAndTasks(group || []);
                    nvsxStr = agg.nvsxStr || '';
                    taskText = (agg.taskLines || []).join('\n');
                }

                const taskLines = (taskText || '').replace(/\r/g,'').split('\n').map(s => s.trim()).filter(Boolean);

                // Consistency check: compare everyone vs final meta (after normalization)
                const finalN = utils.normalizeNvsxInput(nvsxStr);
                const finalT = utils.normalizeTaskText(taskText);

                let inconsistent = false;
                const seenN = new Set();
                const seenT = new Set();
                (group || []).forEach(p => {
                    const pn = utils.normalizeNvsxInput(p?.nvsxNo || '');
                    const pt = utils.normalizeTaskText(p?.taskDesc || '');
                    if (pn) seenN.add(pn);
                    if (pt) seenT.add(pt);
                    if ((pn && pn !== finalN) || (pt && pt !== finalT)) inconsistent = true;
                });

                return {
                    nvsxStr: finalN || '',
                    taskText: finalT || '',
                    taskLines,
                    inconsistent,
                    distinctNvsx: [...seenN],
                    distinctTasks: [...seenT]
                };
            }

        };


        
        const fileUtils = {
            downloadText: (filename, text, mime = 'text/plain;charset=utf-8') => {
                try {
                    const blob = new Blob([text], { type: mime });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                    setTimeout(() => URL.revokeObjectURL(url), 1000);
                } catch (e) {
                    console.error('downloadText failed', e);
                    utils.showToast('Không thể tải file trên trình duyệt này', 'error');
                }
            },
            toCSV: (rows) => {
                const esc = (v) => {
                    const s = String(v ?? '');
                    const needs = /[",\n\r]/.test(s);
                    const out = s.replace(/"/g, '""');
                    return needs ? `"${out}"` : out;
                };
                return rows.map(r => r.map(esc).join(',')).join('\n');
            }
        };

// GLOBAL SAFETY NET (prevent hard crashes; show friendly error)
        window.addEventListener('error', (event) => {
            try {
                const msg = event?.error?.message || event?.message || 'Lỗi không xác định';
                console.error('GlobalError:', event?.error || event);
                utils.showToast('Có lỗi xảy ra: ' + msg, 'error');
            } catch (e) {
                console.error('GlobalErrorHandlerFailed:', e);
            }
        });

        window.addEventListener('unhandledrejection', (event) => {
            try {
                const msg = event?.reason?.message || String(event?.reason || 'Lỗi Promise không xác định');
                console.error('UnhandledRejection:', event?.reason || event);
                utils.showToast('Có lỗi xảy ra: ' + msg, 'error');
            } catch (e) {
                console.error('UnhandledRejectionHandlerFailed:', e);
            }
        });

        const getPublicColl = (name) => db.collection('artifacts').doc(appId).collection('public').doc('data').collection(name);

        const audit = {
            async log(action, details = {}) {
                try {
                    if (!state.currentUser) return;
                    const payload = {
                        action,
                        dateKey: state.currentDateKey || '',
                        user: state.currentUser.email || '',
                        role: state.userRole || '',
                        orgId: state.userOrgId || '',
                        createdAt: new Date(),
                        details: details || {}
                    };
                    await getPublicColl('auditLogs').add(payload);
                } catch (e) {
                    // Do not break main flow if logging fails
                    console.warn('AuditLogFailed:', e);
                }
            }
        };

        const warningCenter = {
            open() {
                try {
                    const dateVal = document.getElementById('planDateInput')?.value || '';
                    document.getElementById('warningDateLabel').innerText = dateVal ? dateVal.split('-').reverse().join('/') : (state.currentDateKey || '--');

                    const list = (state.currentPlanPeople || []).filter(p => Array.isArray(p.warnings) && p.warnings.length > 0);
                    document.getElementById('warningCountLabel').innerText = list.length;

                    // Summary chips by warning type
                    const freq = {};
                    list.forEach(p => (p.warnings || []).forEach(w => { freq[w] = (freq[w] || 0) + 1; }));
                    const summary = Object.keys(freq).sort((a,b)=>freq[b]-freq[a]).map(k => {
                        return `<span class="bg-red-50 text-red-700 border border-red-100 px-2 py-1 rounded-full text-[11px] font-bold">${k}: ${freq[k]}</span>`;
                    }).join('');
                    document.getElementById('warningSummary').innerHTML = summary || '<span class="text-slate-400 italic text-sm">Không có cảnh báo.</span>';

                    warningCenter._all = list;
                    document.getElementById('warningSearch').value = '';
                    warningCenter.render();

                    const modal = document.getElementById('warningModal');
                    modal.classList.remove('hidden');
                    modal.classList.add('flex');

                    // live search
                    document.getElementById('warningSearch').oninput = () => warningCenter.render();
                } catch (e) {
                    console.error(e);
                    utils.showToast('Không mở được danh sách cảnh báo: ' + e.message, 'error');
                }
            },
            close() {
                const modal = document.getElementById('warningModal');
                modal.classList.add('hidden');
                modal.classList.remove('flex');
            },
            render() {
                const q = (document.getElementById('warningSearch').value || '').toLowerCase().trim();
                const rows = (warningCenter._all || []).filter(p => {
                    if (!q) return true;
                    const hay = [
                        p.fullName, p.staffNo, p.orgName, p.destination,
                        ...(p.warnings || [])
                    ].join(' ').toLowerCase();
                    return hay.includes(q);
                });

                const tb = document.getElementById('warningTableBody');
                if (rows.length === 0) {
                    tb.innerHTML = `<tr><td colspan="6" class="px-4 py-6 text-center text-slate-400 italic">Không có kết quả</td></tr>`;
                    return;
                }
                tb.innerHTML = rows.map((p, idx) => `
                    <tr>
                        <td class="px-3 py-2 text-center text-slate-500 font-mono">${idx + 1}</td>
                        <td class="px-3 py-2 font-semibold text-slate-800">${p.fullName || ''}</td>
                        <td class="px-3 py-2 text-center text-slate-700 font-mono">${p.staffNo || ''}</td>
                        <td class="px-3 py-2 text-slate-600">${p.orgName || ''}</td>
                        <td class="px-3 py-2 font-bold text-blue-700">${p.destination || ''}</td>
                        <td class="px-3 py-2 text-red-700">${(p.warnings || []).join('; ')}</td>
                    </tr>
                `).join('');
            },
            exportCSV() {
                try {
                    const rows = (warningCenter._all || []).map(p => [
                        state.currentDateKey,
                        p.destination || '',
                        p.orgName || '',
                        p.fullName || '',
                        p.staffNo || '',
                        (p.warnings || []).join('; ')
                    ]);
                    const header = ['Ngay', 'Gian', 'DonVi', 'HoTen', 'DanhSo', 'CanhBao'];
                    const csv = fileUtils.toCSV([header, ...rows]);
                    fileUtils.downloadText(`canhbao_${state.currentDateKey || 'ngay'}.csv`, csv, 'text/csv;charset=utf-8');
                } catch (e) {
                    console.error(e);
                    utils.showToast('Xuất CSV thất bại: ' + e.message, 'error');
                }
            }
        };

        const auditModal = {
            _cache: [],
            _page: 1,
            _pageSize: 10,
            open() {
                if (state.userRole !== 'dispatcher') return utils.showToast('Chỉ Admin mới xem được nhật ký', 'error');
                const el = document.getElementById('auditModal');
                el.classList.remove('hidden');
                el.classList.add('flex');
                // Wire inputs
                document.getElementById('auditSearch').oninput = () => { auditModal._page = 1; auditModal.render(); };
                document.getElementById('auditActionFilter').onchange = () => { auditModal._page = 1; auditModal.render(); };
                const prevBtn = document.getElementById('auditPrevBtn');
                const nextBtn = document.getElementById('auditNextBtn');
                if (prevBtn) prevBtn.onclick = () => { auditModal._page = Math.max(1, (auditModal._page || 1) - 1); auditModal.render(); };
                if (nextBtn) nextBtn.onclick = () => { auditModal._page = (auditModal._page || 1) + 1; auditModal.render(); };
                if (auditModal._cache.length === 0) auditModal.refresh();
                else auditModal.render();
            },
            close() {
                const el = document.getElementById('auditModal');
                el.classList.add('hidden');
                el.classList.remove('flex');
            },
            async refresh() {
                utils.toggleLoader(true, 'Đang tải nhật ký...');
                try {
                    const snap = await getPublicColl('auditLogs').orderBy('createdAt', 'desc').limit(200).get();
                    auditModal._cache = snap.docs.map(d => ({ id: d.id, ...d.data() }));
                    auditModal._page = 1;
                    auditModal.render();
                } catch (e) {
                    console.error(e);
                    utils.showToast('Không tải được nhật ký: ' + e.message, 'error');
                } finally {
                    utils.toggleLoader(false);
                }
            },
            _fmt(ts) {
                try {
                    if (!ts) return '';
                    const d = ts.seconds ? new Date(ts.seconds * 1000) : new Date(ts);
                    return utils.formatDateTime(d);
                } catch { return ''; }
            },
            _actionLabel(a) {
                const map = {
                    ADD_PERSON: 'Thêm nhân sự',
                    UPDATE_PERSON: 'Sửa nhân sự',
                    DELETE_PERSON: 'Xóa nhân sự',
                    DELETE_ALL: 'Xóa toàn bộ',
                    IMPORT: 'Import',
                    GROUP_EDIT: 'Sửa giàn/NVSX/CV',
                    LOCK_TOGGLE: 'Khóa/Mở ngày',
                    EXPORT_DOCX: 'Export Word'};
                return map[a] || a || '';
            },

render() {
    const q = (document.getElementById('auditSearch').value || '').toLowerCase().trim();
    const act = document.getElementById('auditActionFilter').value || '';
    const all = (auditModal._cache || []).filter(r => {
        if (act && r.action !== act) return false;
        if (!q) return true;
        const hay = [
            r.user,
            r.action,
            r.dateKey,
            JSON.stringify(r.details || {})
        ].join(' ').toLowerCase();
        return hay.includes(q);
    });

    const total = all.length;
    const pageSize = auditModal._pageSize || 10;
    const totalPages = Math.max(1, Math.ceil(total / pageSize));
    auditModal._page = Math.min(Math.max(1, auditModal._page || 1), totalPages);

    const startIdx = (auditModal._page - 1) * pageSize;
    const endIdx = Math.min(startIdx + pageSize, total);
    const pageRows = all.slice(startIdx, endIdx);

    // Update paging UI
    const rangeEl = document.getElementById('auditRange');
    const totalEl = document.getElementById('auditTotal');
    const pageEl = document.getElementById('auditPageLabel');
    const prevBtn = document.getElementById('auditPrevBtn');
    const nextBtn = document.getElementById('auditNextBtn');

    if (rangeEl) rangeEl.textContent = total === 0 ? '0-0' : `${startIdx + 1}-${endIdx}`;
    if (totalEl) totalEl.textContent = String(total);
    if (pageEl) pageEl.textContent = `Trang ${auditModal._page}/${totalPages}`;
    if (prevBtn) {
        prevBtn.disabled = auditModal._page <= 1;
        prevBtn.classList.toggle('opacity-50', prevBtn.disabled);
    }
    if (nextBtn) {
        nextBtn.disabled = auditModal._page >= totalPages;
        nextBtn.classList.toggle('opacity-50', nextBtn.disabled);
    }

    const tb = document.getElementById('auditTableBody');
    if (total === 0) {
        tb.innerHTML = `<tr><td colspan="5" class="px-4 py-6 text-center text-slate-400 italic">Không có dữ liệu</td></tr>`;
        return;
    }

    tb.innerHTML = pageRows.map(r => {
        const details = r.details || {};
        const detailText = auditModal._prettyDetails(details);
        return `
            <tr>
                <td class="px-3 py-2 text-slate-600 font-mono">${auditModal._fmt(r.createdAt)}</td>
                <td class="px-3 py-2 text-slate-700">${r.user || ''}</td>
                <td class="px-3 py-2 font-bold text-blue-700">${auditModal._actionLabel(r.action)}</td>
                <td class="px-3 py-2 text-slate-600 font-mono">${r.dateKey || ''}</td>
                <td class="px-3 py-2 text-slate-600">${detailText}</td>
            </tr>
        `;
    }).join('');
},
            _prettyDetails(details) {
                try {
                    // Keep details short and readable
                    const parts = [];
                    if (details.destination) parts.push(`Giàn: <b>${details.destination}</b>`);
                    if (details.staffNo || details.fullName) parts.push(`NS: <b>${details.fullName || ''}</b> <span class="text-slate-400 font-mono">${details.staffNo || ''}</span>`);
                    if (typeof details.count === 'number') parts.push(`SL: <b>${details.count}</b>`);
                    if (details.status) parts.push(`Trạng thái: <b>${details.status}</b>`);
                    if (details.nvsxNo) parts.push(`NVSX: <b>${details.nvsxNo}</b>`);
                    if (details.reason) parts.push(`Lý do: <span class="text-slate-500">${details.reason}</span>`);
                    if (parts.length > 0) return parts.join(' • ');
                    // fallback (trim)
                    const s = JSON.stringify(details);
                    return s.length > 180 ? s.slice(0, 180) + '…' : s;
                } catch {
                    return '';
                }
            },
            exportCSV() {
                try {
                    const header = ['ThoiGian', 'User', 'HanhDong', 'Ngay', 'ChiTiet'];
                    const rows = (auditModal._cache || []).map(r => [
                        auditModal._fmt(r.createdAt),
                        r.user || '',
                        r.action || '',
                        r.dateKey || '',
                        JSON.stringify(r.details || {})
                    ]);
                    const csv = fileUtils.toCSV([header, ...rows]);
                    fileUtils.downloadText(`nhatky_${(new Date()).toISOString().slice(0,10)}.csv`, csv, 'text/csv;charset=utf-8');
                } catch (e) {
                    console.error(e);
                    utils.showToast('Xuất CSV thất bại: ' + e.message, 'error');
                }
            }
        };

        const masterDataManager = {
            saveToMaster: (people) => {
                const batch = db.batch();
                let count = 0;
                people.forEach(p => {
                    if (p.staffNo) {
                        const ref = getPublicColl('masterData').doc(p.staffNo);
                        const data = {
                            staffNo: p.staffNo, fullName: p.fullName, title: p.title, orgName: p.orgName,
                            dob: p.dob, pob: p.pob, phone: p.phone, idNo: p.idNo,
                            idIssueDate: p.idIssueDate, idExpiryDate: p.idExpiryDate,
                            weightKg: p.weightKg, luggageKg: p.luggageKg, cargoKg: p.cargoKg,
                            lastUpdated: new Date()
                        };
                        batch.set(ref, data, { merge: true });
                        count++;
                    }
                });
                if (count > 0) batch.commit().catch(console.error);
            },
            autoFill: async (staffNo) => {
                if (!staffNo) return;
                utils.toggleLoader(true, "Đang tra cứu...");
                try {
                    const doc = await getPublicColl('masterData').doc(staffNo.trim()).get();
                    if (doc.exists) {
                        const d = doc.data();
                        document.getElementById('edit_fullName').value = d.fullName || '';
                        document.getElementById('edit_title').value = d.title || '';
                        
                        // CHANGE: Only populate orgName if user is Dispatcher/Admin
                        if (state.userRole === 'dispatcher') {
                            document.getElementById('edit_orgName').value = d.orgName || '';
                        }
                        
                        document.getElementById('edit_dob').value = d.dob || '';
                        document.getElementById('edit_pob').value = d.pob || '';
                        document.getElementById('edit_phone').value = d.phone || '';
                        document.getElementById('edit_idNo').value = d.idNo || '';
                        document.getElementById('edit_idIssueDate').value = d.idIssueDate || '';
                        document.getElementById('edit_idExpiryDate').value = d.idExpiryDate || '';
                        document.getElementById('edit_weightKg').value = d.weightKg || '';
                        document.getElementById('edit_luggageKg').value = d.luggageKg || '';
                        document.getElementById('edit_cargoKg').value = d.cargoKg || '';
                        utils.showToast("Đã tìm thấy thông tin nhân sự", "success");
                        // NEW: Show destinations when user is found/searching
                        document.getElementById('advEdit').classList.remove('hidden');
                        editModal.renderDestinationSuggestions();
                    } else {
                        utils.showToast("Không tìm thấy trong hồ sơ", "error");
                    }
                } catch (e) { console.error(e); utils.showToast("Lỗi: " + e.message, "error"); } finally { utils.toggleLoader(false); }
            }
        };

        const conflictChecker = {
            check: (newPeople, currentPeople, isEditId = null) => {
                const conflicts = [];
                newPeople.forEach(np => {
                    const found = currentPeople.find(cp => {
                        if (isEditId && cp.id === isEditId) return false;
                        return (np.staffNo && cp.staffNo === np.staffNo) || (np.idNo && cp.idNo === np.idNo);
                    });
                    if (found) {
                        conflicts.push({
                            new: np,
                            current: found,
                            msg: `Nhân sự ${np.fullName} (DS: ${np.staffNo}) đã có tên trong danh sách đi ${found.destination}`
                        });
                    }
                });
                return conflicts;
            }
        };

        const reportManager = {
            selectedDates: [],
            mode: 'days', // 'days', 'month', 'year'
            
            open: () => {
                document.getElementById('reportModal').classList.remove('hidden');
                document.getElementById('reportModal').classList.add('flex');
                
                // Init View
                reportManager.setMode('days');
                
                // Populate Year Select
                const yearSelect = document.getElementById('rptYearInput');
                yearSelect.innerHTML = '';
                const currentYear = new Date().getFullYear();
                for (let y = currentYear; y >= 2024; y--) {
                    const opt = document.createElement('option');
                    opt.value = y;
                    opt.text = `Năm ${y}`;
                    yearSelect.appendChild(opt);
                }
                
                // Pre-fill Today
                const todayStr = document.getElementById('planDateInput').value;
                if(todayStr) {
                     document.getElementById('rptDateInput').value = todayStr;
                     document.getElementById('rptMonthInput').value = todayStr.substring(0, 7);
                }
            },
            
            close: () => {
                document.getElementById('reportModal').classList.add('hidden');
                document.getElementById('reportModal').classList.remove('flex');
            },
            
            setMode: (mode) => {
                reportManager.mode = mode;
                // UI Toggle
                ['days', 'month', 'year'].forEach(m => {
                    const btn = document.getElementById(`rptMode${m.charAt(0).toUpperCase() + m.slice(1)}`);
                    if(m === mode) {
                        btn.className = "flex-1 py-1.5 text-xs font-bold rounded-md transition bg-white shadow text-purple-700";
                    } else {
                        btn.className = "flex-1 py-1.5 text-xs font-bold rounded-md transition text-slate-500 hover:text-slate-800";
                    }
                });
                
                document.getElementById('rptDaysView').classList.toggle('hidden', mode !== 'days');
                document.getElementById('rptMonthView').classList.toggle('hidden', mode !== 'month');
                document.getElementById('rptYearView').classList.toggle('hidden', mode !== 'year');
            },
            
            addDate: () => {
                const val = document.getElementById('rptDateInput').value;
                if(val && !reportManager.selectedDates.includes(val)) {
                    reportManager.selectedDates.push(val);
                    reportManager.renderDateList();
                }
            },
            
            removeDate: (val) => {
                reportManager.selectedDates = reportManager.selectedDates.filter(d => d !== val);
                reportManager.renderDateList();
            },
            
            renderDateList: () => {
                const ul = document.getElementById('rptDateList');
                ul.innerHTML = reportManager.selectedDates.map(d => {
                    const [y,m,day] = d.split('-');
                    return `<li class="flex justify-between items-center bg-white p-1.5 rounded border border-slate-200 text-xs">
                        <span class="font-bold text-slate-700">${day}/${m}/${y}</span>
                        <button onclick="reportManager.removeDate('${d}')" class="text-red-500 hover:text-red-700 font-bold px-2">&times;</button>
                    </li>`;
                }).join('');
            },
            
            generateDateRange: () => {
                if (reportManager.mode === 'days') return reportManager.selectedDates;
                
                let start, end;
                if (reportManager.mode === 'month') {
                    const mVal = document.getElementById('rptMonthInput').value; // YYYY-MM
                    if(!mVal) return [];
                    const [y, m] = mVal.split('-');
                    start = new Date(y, m-1, 1);
                    end = new Date(y, m, 0);
                } else if (reportManager.mode === 'year') {
                    const yVal = document.getElementById('rptYearInput').value;
                    start = new Date(yVal, 0, 1);
                    end = new Date(yVal, 11, 31);
                }
                
                const dates = [];
                for (let d = start; d <= end; d.setDate(d.getDate() + 1)) {
                    const mm = String(d.getMonth() + 1).padStart(2, '0');
                    const dd = String(d.getDate()).padStart(2, '0');
                    const yyyy = d.getFullYear();
                    dates.push(`${yyyy}-${mm}-${dd}`);
                }
                return dates;
            },
            
            processReport: async () => {
                const dates = reportManager.generateDateRange();
                if (dates.length === 0) return utils.showToast('Vui lòng chọn thời gian', 'error');
                
                // Show Progress
                const pContainer = document.getElementById('rptProgress');
                const pBar = document.getElementById('rptProgressBar');
                const pText = document.getElementById('rptProgressPct');
                
                pContainer.classList.remove('hidden');
                
                let allPeople = [];
                const chunkSize = 5; // Batch 5 days per request to avoid hitting limits too hard
                
                try {
                    for (let i = 0; i < dates.length; i += chunkSize) {
                        const batchDates = dates.slice(i, i + chunkSize);
                        const promises = batchDates.map(async dateStr => {
                            const dateKey = dateStr.replace(/-/g, '');
                            const snap = await getPublicColl('dailyPlans').doc(dateKey).collection('people').get();
                            // Inject Date Info and sortable date
                            const [y, m, d] = dateStr.split('-');
                            return snap.docs.map(doc => ({
                                ...doc.data(), 
                                _reportDate: `${d}/${m}/${y}`,
                                _sortDate: dateStr // YYYY-MM-DD for sorting
                            }));
                        });
                        
                        const batchResults = await Promise.all(promises);
                        allPeople = allPeople.concat(batchResults.flat());
                        
                        // Update Progress
                        const pct = Math.round(((i + batchDates.length) / dates.length) * 100);
                        pBar.style.width = `${pct}%`;
                        pText.innerText = `${pct}%`;
                    }
                    
                    if(allPeople.length === 0) {
                         utils.showToast("Không có dữ liệu trong khoảng thời gian này", "warning");
                         pContainer.classList.add('hidden');
                         return;
                    }
                    
                    // Generate Excel
                    reportManager.exportToExcel(allPeople);
                    
                } catch(e) {
                    console.error(e);
                    utils.showToast("Lỗi tải dữ liệu: " + e.message, "error");
                } finally {
                     pContainer.classList.add('hidden');
                     pBar.style.width = '0%';
                }
            },
            
            exportToExcel: (data) => {
                if (typeof XLSX === 'undefined') {
                    utils.showToast('Không tải được thư viện Excel (XLSX). Hãy kiểm tra kết nối internet/CDN.', 'error');
                    return;
                }

                // Sort by Sortable Date, then Org, then Name
                data.sort((a,b) => {
                    return a._sortDate.localeCompare(b._sortDate) || a.orgName.localeCompare(b.orgName);
                });
                
                const exportData = data.map((p, idx) => ({
                    "STT": idx + 1,
                    "Ngày đăng ký": p._reportDate,
                    "Đơn vị": p.orgName,
                    "Danh số": p.staffNo,
                    "Họ và tên": p.fullName,
                    "Chức danh": p.title,
                    "Nơi đến": p.destination,
                    "Trọng lượng (Kg)": p.weightKg,
                    "Hành lý (Kg)": p.luggageKg,
                    "Hàng hóa (Kg)": p.cargoKg,
                    "Ghi chú": p.rowNote
                }));
                
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.json_to_sheet(exportData);
                
                // Auto-width
                const wscols = [
                    {wch: 5}, {wch: 12}, {wch: 20}, {wch: 10}, {wch: 25}, 
                    {wch: 20}, {wch: 15}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 20}
                ];
                ws['!cols'] = wscols;

                XLSX.utils.book_append_sheet(wb, ws, "Báo Cáo");
                XLSX.writeFile(wb, `BaoCao_XNXL_${new Date().getTime()}.xlsx`);
                
                utils.showToast("Đã tải xuống file Excel", "success");
                reportManager.close();
            }
        };

        const userModal = {
            open: (userId = null) => {
                const titleEl = document.getElementById('userModalTitle');
                if (userId) {
                    // Edit Mode
                    const user = state.users.find(u => u.id === userId);
                    if (!user) return;
                    document.getElementById('u_id').value = userId;
                    document.getElementById('u_email').value = user.email;
                    document.getElementById('u_email').disabled = true;
                    document.getElementById('u_email').classList.add('bg-gray-100');
                    document.getElementById('u_passwordWrapper').classList.add('hidden');
                    document.getElementById('u_name').value = user.displayName || '';
                    document.getElementById('u_role').value = user.role;
                    titleEl.innerText = "Sửa thông tin User";
                } else {
                    // Create Mode
                    document.getElementById('u_id').value = '';
                    document.getElementById('u_email').value = '';
                    document.getElementById('u_email').disabled = false;
                    document.getElementById('u_email').classList.remove('bg-gray-100');
                    document.getElementById('u_password').value = '';
                    document.getElementById('u_passwordWrapper').classList.remove('hidden');
                    document.getElementById('u_name').value = '';
                    document.getElementById('u_role').value = 'workshop';
                    titleEl.innerText = "Thêm User Mới";
                }
                userModal.toggleOrgSelect();
                // Set org based on edit data or default
                if (userId) {
                    const user = state.users.find(u => u.id === userId);
                    if (user && user.role !== 'dispatcher') {
                         document.getElementById('u_orgId').value = user.orgId;
                    }
                }
                document.getElementById('userModal').classList.remove('hidden'); 
                document.getElementById('userModal').classList.add('flex');
            },
            close: () => { 
                document.getElementById('userModal').classList.add('hidden'); 
                document.getElementById('userModal').classList.remove('flex'); 
            },
            toggleOrgSelect: () => {
                const role = document.getElementById('u_role').value;
                document.getElementById('u_orgWrapper').classList.toggle('hidden', role === 'dispatcher');
            },
            save: () => {
                const uid = document.getElementById('u_id').value;
                if (uid) {
                    userManager.updateUser(uid);
                } else {
                    userManager.createUser();
                }
            }
        };

        const userManager = {
            loadUsers: async () => {
                const snap = await db.collection('artifacts').doc(appId).collection('users').get();
                state.users = snap.docs.map(doc => ({id: doc.id, ...doc.data()}));
                document.getElementById('userListBody').innerHTML = state.users.map(u => {
                    const roleBadge = u.role === 'dispatcher' ? '<span class="bg-red-100 text-red-800 text-xs px-2 py-1 rounded-full font-bold">Admin</span>' : '<span class="bg-blue-100 text-blue-800 text-xs px-2 py-1 rounded-full font-bold">User</span>';
                    return `<tr class="hover:bg-slate-50">
                        <td class="px-4 py-2">${u.email}</td>
                        <td class="px-4 py-2 font-medium">${u.displayName || '-'}</td>
                        <td class="px-4 py-2">${roleBadge}</td>
                        <td class="px-4 py-2 text-slate-500">${u.orgId || 'XNXL'}</td>
                        <td class="px-4 py-2 text-right">
                            <button onclick="userModal.open('${u.id}')" class="text-slate-500 hover:text-blue-600 mr-2" title="Sửa"><i class="fas fa-edit"></i></button>
                            <button onclick="userManager.sendResetEmail('${u.email}')" class="text-blue-600 hover:text-blue-800 text-xs font-bold mr-2">Reset Pass</button>
                            <button onclick="userManager.deleteUserDoc('${u.id}')" class="text-red-600 hover:text-red-800"><i class="fas fa-trash"></i></button>
                        </td>
                    </tr>`;
                }).join('');
            },
            createUser: async () => {
                const email = document.getElementById('u_email').value.trim();
                const password = document.getElementById('u_password').value.trim();
                const name = document.getElementById('u_name').value.trim();
                const role = document.getElementById('u_role').value;
                const orgId = role === 'dispatcher' ? 'XNXL' : document.getElementById('u_orgId').value;
                if(!email || !password || !name) return utils.showToast('Vui lòng điền đủ thông tin', 'error');
                utils.toggleLoader(true, 'Đang tạo user...');
                try {
                    const secondaryApp = firebase.initializeApp(firebaseConfig, "Secondary");
                    const userCred = await secondaryApp.auth().createUserWithEmailAndPassword(email, password);
                    await db.collection('artifacts').doc(appId).collection('users').doc(userCred.user.uid).set({ email, role, orgId, displayName: name, createdAt: new Date() });
                    await secondaryApp.delete();
                    utils.showToast(`Đã tạo user ${email}`, 'success'); userModal.close(); userManager.loadUsers();
                } catch (e) { utils.showToast(e.message, 'error'); } finally { utils.toggleLoader(false); }
            },
            updateUser: async (uid) => {
                const name = document.getElementById('u_name').value.trim();
                const role = document.getElementById('u_role').value;
                const orgId = role === 'dispatcher' ? 'XNXL' : document.getElementById('u_orgId').value;
                
                utils.toggleLoader(true, 'Đang cập nhật...');
                try {
                    await db.collection('artifacts').doc(appId).collection('users').doc(uid).update({
                        displayName: name,
                        role: role,
                        orgId: orgId
                    });
                    utils.showToast('Đã cập nhật thông tin user', 'success');
                    userModal.close();
                    userManager.loadUsers();
                } catch(e) {
                    utils.showToast(e.message, 'error');
                } finally {
                    utils.toggleLoader(false);
                }
            },
            sendResetEmail: async (email) => { if(confirm(`Gửi email đặt lại mật khẩu cho ${email}?`)) auth.sendPasswordResetEmail(email).then(() => utils.showToast('Đã gửi email reset pass', 'success')).catch(e => utils.showToast(e.message, 'error')); },
            deleteUserDoc: async (uid) => { if(confirm('Xóa user này?')) db.collection('artifacts').doc(appId).collection('users').doc(uid).delete().then(() => { utils.showToast('Đã xóa profile', 'success'); userManager.loadUsers(); }).catch(e => utils.showToast(e.message, 'error')); }
        };

        const importModal = {
            parsedData: null,
            getSelectedOrgName: () => {
                try {
                    if (state.userRole === 'dispatcher') {
                        const sel = document.getElementById('importOrgSelect');
                        return sel && sel.selectedIndex >= 0 ? (sel.options[sel.selectedIndex].text || '') : '';
                    }
                    const found = state.orgs.find(o => o.id === state.userOrgId);
                    return found ? found.name : (state.userOrgId || '');
                } catch (e) { return ''; }
            },
            open: () => {
                document.getElementById('excelFile').value = ''; document.getElementById('fileNameDisplay').innerText = ''; document.getElementById('prevDateInput').value=''; document.getElementById('prevDestInput').value=''; document.getElementById('prevNVSXInput').value=''; document.getElementById('prevTaskInput').value='';
                document.getElementById('parseMeta').classList.add('hidden'); document.getElementById('parsePeopleWrap').classList.add('hidden'); document.getElementById('btnConfirmImport').disabled = true;
                document.getElementById('importModal').classList.remove('hidden'); document.getElementById('importModal').classList.add('flex'); const sel = document.getElementById('importOrgSelect'); if(sel && !sel._boundChange){ sel.addEventListener('change', ()=>{ if(importModal.parsedData){ importModal.renderPeoplePreview(importModal.parsedData.people);} }); sel._boundChange=true; } const destIn=document.getElementById('prevDestInput'); if(destIn && !destIn._boundInput){ destIn.addEventListener('input', ()=>destIn.classList.remove('border-red-400')); destIn._boundInput=true; }
                const dateIn=document.getElementById('prevDateInput'); if(dateIn && !dateIn._boundInput){ dateIn.addEventListener('input', ()=>{ const sel=document.getElementById('planDateInput')?.value; if(sel){ const [y,m,d]=sel.split('-'); const expect=`${d}/${m}/${y}`; const w=document.getElementById('prevWarning'); if(w){ if(dateIn.value.trim() && dateIn.value.trim()!==expect){ w.classList.remove('hidden'); w.title=`File: ${dateIn.value} khác Ngày đang chọn`; } else { w.classList.add('hidden'); } } } }); dateIn._boundInput=true; }
            },
            close: () => { document.getElementById('importModal').classList.add('hidden'); document.getElementById('importModal').classList.remove('flex'); },
            handleFile: async (e) => {
                const file = e.target.files[0]; if(!file) return;
                document.getElementById('fileNameDisplay').innerText = file.name; utils.toggleLoader(true, 'Đang phân tích file Excel...');
                try {
                    const data = await importModal.parseExcel(file); importModal.parsedData = data;
                    document.getElementById('prevDateInput').value = data.date || ''; document.getElementById('prevDestInput').value = data.destination || ''; document.getElementById('prevDestInput').classList.toggle('border-red-400', !(data.destination||'').trim());
                    document.getElementById('prevNVSXInput').value = data.nvsx || ''; document.getElementById('prevTaskInput').value = data.task || '';
                    document.getElementById('prevCount').innerText = data.people.length;
                    importModal.renderPeoplePreview(data.people);

                    const selectedDate = document.getElementById('planDateInput').value; const [y, m, d] = selectedDate.split('-');
                    if(data.date && data.date !== `${d}/${m}/${y}`) { document.getElementById('prevWarning').classList.remove('hidden'); document.getElementById('prevWarning').title = `File: ${data.date} khác Ngày đang chọn`; }
                    else document.getElementById('prevWarning').classList.add('hidden');
                    document.getElementById('parseMeta').classList.remove('hidden'); document.getElementById('parsePeopleWrap').classList.remove('hidden'); document.getElementById('btnConfirmImport').disabled = data.people.length === 0;
                } catch (err) { utils.showToast('Lỗi đọc file: ' + err.message, 'error'); } finally { utils.toggleLoader(false); }
            },
            
            renderPeoplePreview: (people) => {
                const tbody = document.getElementById('prevPeopleBody');
                const more = document.getElementById('prevPeopleMore');
                if (!tbody) return;
                const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
                const max = 300;
                const slice = (people || []).slice(0, max);
                tbody.innerHTML = slice.map((p, i) => `
                    <tr>
                        <td class="px-3 py-2 text-slate-500">${i + 1}</td>
                        <td class="px-3 py-2 font-semibold text-slate-800">${esc(p.staffNo || '')}</td>
                        <td class="px-3 py-2 text-slate-800">${esc(p.fullName || '')}</td>
                        <td class="px-3 py-2 text-slate-600">${esc(p.title || '')}</td>
                        <td class="px-3 py-2 text-slate-500">${esc(importModal.getSelectedOrgName() || p.orgName || '')}</td>
                    </tr>
                `).join('');
                if (more) {
                    if ((people || []).length > max) {
                        more.classList.remove('hidden');
                        more.textContent = `Hiển thị ${max}/${people.length} người (danh sách quá dài). Vẫn có thể Import đầy đủ.`;
                    } else {
                        more.classList.add('hidden');
                        more.textContent = '';
                    }
                }
            },parseExcel: (file) => {
                return new Promise((resolve, reject) => {
                    if (typeof XLSX === 'undefined') {
                        reject(new Error('Không tải được thư viện Excel (XLSX). Hãy kiểm tra kết nối internet/CDN.'));
                        return;
                    }

                    const reader = new FileReader();
                    reader.onerror = () => reject(new Error('Không đọc được file Excel. Vui lòng thử lại hoặc chọn file khác.'));
                    reader.onload = (e) => {
                        try {
                            const workbook = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
                            const sheet = workbook.Sheets[workbook.SheetNames[0]];
                            const rows = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
                            let date = '', destination = '', nvsx = '', task = '', headerRowIndex = -1;

                            // Prefer fixed cells for "Mẫu ĐK đi biển - Form mới" when present
                            const __cellVal = (addr) => {
                                try {
                                    const c = sheet && sheet[addr];
                                    if (!c) return '';
                                    return (c.v !== undefined ? c.v : (c.w !== undefined ? c.w : ''));
                                } catch (e) { return ''; }
                            };
                            const __d6 = __cellVal('D6');
                            if (__d6 && !date) date = utils.processExcelDate(__d6) || '';
                            const __d7 = __cellVal('D7');
                            const __e7 = __cellVal('E7');
                            if (!destination && (__d7 || __e7)) {
                                destination = [__d7, __e7].map(v => (v ?? '').toString().trim()).filter(Boolean).join(' ').replace(/\s+/g,' ').trim();
                            }
                            const __d8 = __cellVal('D8');
                            if (__d8 && !nvsx) nvsx = __d8.toString().trim();
                            const __d9 = __cellVal('D9');
                            if (__d9 && !task) task = __d9.toString().replace(/\r/g,'').trim();

                            
                            // Extract Meta (robust, supports Form cũ & Form mới)
                            const __normCell = (v) => {
                                try {
                                    if (v === undefined || v === null) return '';
                                    return v.toString().trim().toLowerCase()
                                        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
                                        .replace(/\s+/g, ' ');
                                } catch (e) { return ''; }
                            };
                            const __isEmpty = (v) => v === undefined || v === null || (typeof v === 'string' && v.trim() === '');
                            const __skipTokens = new Set(['gian','giàn','tau','tàu','t\\u00e0u','form','rev','don vi','\\u0111\\u01a1n vi','don vị','\\u0111on vi']);
                            const __looksLikeDest = (v) => {
                                const s = (v ?? '').toString().trim();
                                if (!s) return false;
                                const up = s.toUpperCase();
                                // Common rig/ship codes: CTK-3, MSP-9, BK-14, etc.
                                return /^(CTK|MSP|BK|BKP|CT|TB|TD|GT|GIA?N|TAU|T\\u00c0U)[\\s\\-]*\\d+/i.test(up) || /\\b(CTK|MSP|BK)[\\s\\-]*\\d+\\b/i.test(up);
                            };
                            const __findRightValue = (ri, ci) => {
                                const row = rows[ri] || [];
                                const parts = [];
                                for (let k = ci + 1; k < row.length; k++) {
                                    const v = row[k];
                                    if (__isEmpty(v)) continue;
                                    const ns = __normCell(v);
                                    if (!ns) continue;
                                    if (__skipTokens.has(ns)) continue;
                                    if (typeof v === 'string' && __normCell(v).endsWith(':')) continue;
                                    parts.push(v.toString().trim());
                                    // if next cell looks like something else, stop; otherwise allow join of 2 cells (e.g. "GIÀN" + "CTK-3")
                                    if (parts.length >= 1) break;
                                }
                                if (parts.length) return parts.join(' ').replace(/\s+/g,' ').trim();
                                return null;
                            };
                            const __findByLabel = (keywords) => {
                                for (let i = 0; i < Math.min(rows.length, 80); i++) {
                                    const row = rows[i] || [];
                                    for (let j = 0; j < row.length; j++) {
                                        const cell = row[j];
                                        const sc = __normCell(cell);
                                        if (!sc) continue;
                                        if (keywords.some(k => sc.includes(k))) {
                                            const right = __findRightValue(i, j);
                                            if (right) return right;
                                            // If label & value are in same cell (rare)
                                            if (typeof cell === 'string') {
                                                const pos = cell.indexOf(':');
                                                if (pos > -1) {
                                                    const after = cell.slice(pos + 1).trim();
                                                    if (after) return after;
                                                }
                                            }
                                            // Fallback: search nearby cells (right/down within 2 rows)
                                            for (let di = 0; di <= 2; di++) {
                                                const rr = rows[i + di] || [];
                                                for (let dj = 1; dj <= 4; dj++) {
                                                    const v = rr[j + dj];
                                                    if (__isEmpty(v)) continue;
                                                    const ns = __normCell(v);
                                                    if (__skipTokens.has(ns)) continue;
                                                    return v;
                                                }
                                            }
                                        }
                                    }
                                }
                                return null;
                            };

                            // Date
                            if (!date) {
                            const __dateVal = __findByLabel(['dang ky di bien ngay','dang ky ngay','ngay dk','ngay dang ky','ngay di','ngay khoi hanh']);
                            if (__dateVal) {
                                date = utils.processExcelDate(__dateVal) || '';
                                if (!date && typeof __dateVal === 'string') {
                                    const dm = __dateVal.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
                                    if (dm) {
                                        const y = (dm[3].length === 2) ? ('20' + dm[3]) : dm[3];
                                        date = `${dm[1].padStart(2,'0')}/${dm[2].padStart(2,'0')}/${y}`;
                                    }
                                }
                            }
                            
                            }
// Destination
                            if (!destination) {
                            let __destVal = __findByLabel(['noi den','di den','noi den gian','noi den tau','noi den:']);
                            // Special case: destination may appear after a "GIÀN"/"TÀU" token
                            if (__destVal && __skipTokens.has(__normCell(__destVal))) __destVal = null;
                            if (!__destVal) {
                                // Fallback: scan first 20 rows for a likely destination code
                                for (let i = 0; i < Math.min(rows.length, 20) && !__destVal; i++) {
                                    for (let j = 0; j < (rows[i] || []).length; j++) {
                                        const v = rows[i][j];
                                        if (__isEmpty(v)) continue;
                                        if (typeof v === 'string' && __looksLikeDest(v)) { __destVal = v; break; }
                                    }
                                }
                            }
                            if (__destVal) destination = __destVal.toString().trim().replace(/\s+/g,' ');

                            
                            }
// NVSX
                            if (!nvsx) {
                            const __nvsxVal = __findByLabel(['nhiem vu san xuat so','nhiem vu sx so','nhiem vu san xuat','nvsx']);
                            if (__nvsxVal) {
                                const val = __nvsxVal.toString().trim();
                                const m = val.match(/\d{3}-\d{2}(?:bs\d+)?/gi);
                                nvsx = m ? [...new Set(m.map(x => x.toUpperCase()))].join('; ') : val.replace(/^.*:\s*/, '');
                            } else {
                                // Fallback: scan first 30 rows for NVSX tokens
                                const found = [];
                                for (let i = 0; i < Math.min(rows.length, 30); i++) {
                                    for (let j = 0; j < (rows[i] || []).length; j++) {
                                        const v = rows[i][j];
                                        if (__isEmpty(v)) continue;
                                        const str = v.toString();
                                        const m = str.match(/\d{3}-\d{2}(?:bs\d+)?/gi);
                                        if (m) found.push(...m);
                                    }
                                }
                                const uniq = [...new Set(found.map(x => x.toUpperCase()))];
                                if (uniq.length) nvsx = uniq.join('; ');
                            }

                            
                            }
// Task / Work content
                            if (!task) {
                            const __taskVal = __findByLabel(['cong viec thuc hien','noi dung cong viec','noi dung nvsx','cong viec']);
                            if (__taskVal) {
                                let t = __taskVal.toString().replace(/\r/g,'').trim();
                                // Remove a leading label only if it actually starts with "công việc..."
                                t = t.replace(/^\s*(công việc thực hiện|công việc)\s*:?\s*/i, '');
                                task = t;
                            } else {
                                // Fallback: choose the longest multi-line cell containing NVSX-like lines
                                let best = '';
                                for (let i = 0; i < Math.min(rows.length, 35); i++) {
                                    for (let j = 0; j < (rows[i] || []).length; j++) {
                                        const v = rows[i][j];
                                        if (__isEmpty(v)) continue;
                                        const str = v.toString().replace(/\r/g,'');
                                        if (str.includes('\n') && /\d{3}-\d{2}.*:/i.test(str)) {
                                            if (str.length > best.length) best = str;
                                        }
                                    }
                                }
                                if (best) task = best.trim();
                            }

                            
                            }
// Find Header (robust)
                            const __normHeader = (v) => ((v === undefined || v === null) ? '' : v).toString().trim().toLowerCase()
                                .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
                                .replace(/\s+/g, ' ');
                            const __rowHas = (cells, keywords) => cells.some(c => keywords.some(k => c.includes(k)));
                            const __rowScore = (row) => {
                                const cells = (row || []).map(__normHeader).filter(Boolean);
                                if (cells.length === 0) return 0;
                                let score = 0;
                                if (__rowHas(cells, ['ho va ten','ho ten','ten','full name','fullname','name'])) score++;
                                if (__rowHas(cells, ['danh so','msnv','ma nv','ma so','staff no','staffno','employee','emp no'])) score++;
                                if (__rowHas(cells, ['don vi','xuong','org','department','bo phan'])) score++;
                                if (__rowHas(cells, ['chuc danh','chuc vu','position','title'])) score++;
                                if (__rowHas(cells, ['stt','no.','so tt'])) score++;
                                return score;
                            };

                            for (let i = 0; i < Math.min(rows.length, 80); i++) {
                                const sc = __rowScore(rows[i]);
                                // Require at least "Họ và tên" + one more important column
                                if (sc >= 2 && __rowHas((rows[i] || []).map(__normHeader), ['ho va ten','ho ten'])) {
                                    headerRowIndex = i;
                                    break;
                                }
                            }
                            if (headerRowIndex === -1) {
                                return reject(new Error("Không tìm thấy dòng tiêu đề bảng. Hãy kiểm tra file Excel có các cột như: 'Họ và tên', 'Danh số/MSNV', 'Đơn vị/Xưởng'..."));
                            }

                            const headers = rows[headerRowIndex].map(h => h.toString().toLowerCase().trim());
                            const colMap = {
                                staffNo: headers.findIndex(h => h.includes('danh số') || h.includes('msnv')),
                                fullName: headers.findIndex(h => h.includes('họ và tên')),
                                title: headers.findIndex(h => h.includes('chức danh')),
                                org: headers.findIndex(h => h.includes('đơn vị') || h.includes('xưởng')),
                                dob: headers.findIndex(h => h.includes('sinh') && !h.includes('nơi')),
                                pob: headers.findIndex(h => h.includes('nơi sinh')),
                                phone: headers.findIndex(h => h.includes('sđt')),
                                idNo: headers.findIndex(h => h.includes('cccd') || h.includes('cmnd')),
                                idIssue: headers.findIndex(h => h.includes('ngày cấp')),
                                idExpiry: headers.findIndex(h => h.includes('hết hạn')),
                                weight: headers.findIndex(h => h.includes('kg') || h.includes('cân nặng')),
                                luggage: headers.findIndex(h => h.includes('hành lý')),
                                cargo: headers.findIndex(h => h.includes('hàng hóa')),
                                note: headers.findIndex(h => h.includes('ghi chú'))
                            };

                            const people = [];
                            for(let i = headerRowIndex + 1; i < rows.length; i++) {
                                const row = rows[i]; if(!row || !row[colMap.fullName]) continue;
                                const name = row[colMap.fullName].toString().trim();
                                if(!name || name.toLowerCase().includes('tổng cộng')) continue;
                                people.push({
                                    fullName: name,
                                    staffNo: colMap.staffNo > -1 ? utils.cleanString(row[colMap.staffNo]) : '',
                                    title: colMap.title > -1 ? utils.cleanString(row[colMap.title]) : '',
                                    orgName: colMap.org > -1 ? utils.cleanString(row[colMap.org]) : '',
                                    dob: colMap.dob > -1 ? utils.processExcelDate(row[colMap.dob]) : '',
                                    pob: colMap.pob > -1 ? utils.cleanString(row[colMap.pob]) : '',
                                    phone: colMap.phone > -1 ? utils.cleanString(row[colMap.phone]) : '',
                                    idNo: colMap.idNo > -1 ? utils.cleanString(row[colMap.idNo]) : '',
                                    idIssueDate: colMap.idIssue > -1 ? utils.processExcelDate(row[colMap.idIssue]) : '',
                                    idExpiryDate: colMap.idExpiry > -1 ? utils.processExcelDate(row[colMap.idExpiry]) : '',
                                    weightKg: colMap.weight > -1 ? (parseFloat(row[colMap.weight]) || 0) : 0,
                                    luggageKg: colMap.luggage > -1 ? (parseFloat(row[colMap.luggage]) || 0) : 0,
                                    cargoKg: colMap.cargo > -1 ? (parseFloat(row[colMap.cargo]) || 0) : 0,
                                    rowNote: colMap.note > -1 ? utils.cleanString(row[colMap.note]) : ''
                                });
                            }
                            resolve({ date, destination, nvsx, task, people });
                        } catch (err) { reject(err); }
                    };
                    reader.readAsArrayBuffer(file);
                });
            },
            processImport: async () => {
                if(!importModal.parsedData) return;
                
                const people = importModal.parsedData.people || [];
                const date = (document.getElementById('prevDateInput')?.value || '').trim();
                const destination = (document.getElementById('prevDestInput')?.value || '').trim();
                const nvsx = (document.getElementById('prevNVSXInput')?.value || '').trim();
                const task = (document.getElementById('prevTaskInput')?.value || '').replace(/\r/g,'').trim();
                if(!destination){ utils.showToast('Bắt buộc nhập Nơi đến (Giàn/Tàu) trước khi Import.', 'error'); document.getElementById('prevDestInput')?.focus(); return; }
                
                
                // 1. Identify Conflicts
                const conflictResults = conflictChecker.check(people, state.currentPlanPeople);
                const duplicateStaffNos = conflictResults.map(c => c.new.staffNo);
                
                // 2. Filter New People Only
                const newPeople = people.filter(p => !duplicateStaffNos.includes(p.staffNo));
                
                // 3. Always Show Result Modal (Success + Conflicts)
                const tbodySuccess = document.getElementById('successListBody');
                const tbodyConflict = document.getElementById('conflictListBody');
                
                // Render Success List
                tbodySuccess.innerHTML = newPeople.length > 0 ? newPeople.map(p => `
                    <tr>
                        <td class="px-4 py-2 font-medium text-emerald-800">${p.fullName}</td>
                        <td class="px-4 py-2 font-mono text-emerald-600">${p.staffNo}</td>
                        <td class="px-4 py-2 text-emerald-600">${p.title}</td>
                        <td class="px-4 py-2 text-emerald-600">${p.orgName || (state.userRole !== 'dispatcher' ? state.userOrgId : '')}</td>
                    </tr>
                `).join('') : `<tr><td colspan="4" class="px-4 py-4 text-center text-slate-400 italic">Không có nhân sự mới</td></tr>`;
                document.getElementById('successCount').innerText = newPeople.length;

                // Render Conflict List
                tbodyConflict.innerHTML = conflictResults.length > 0 ? conflictResults.map(c => {
                    const t = c.current.createdAt ? c.current.createdAt.seconds*1000 : Date.now();
                    const u = (c.current.importedBy || c.current.updatedBy || 'unknown').split('@')[0];
                    return `
                        <tr>
                            <td class="px-4 py-2 font-medium text-slate-800">${c.new.fullName}</td>
                            <td class="px-4 py-2 font-mono text-slate-600">${c.new.staffNo}</td>
                            <td class="px-4 py-2 font-bold text-blue-600">${c.current.destination}</td>
                            <td class="px-4 py-2 text-slate-500">${u}</td>
                            <td class="px-4 py-2 text-slate-400 text-[11px]">${utils.formatDateTime(new Date(t))}</td>
                        </tr>
                    `;
                }).join('') : `<tr><td colspan="5" class="px-4 py-4 text-center text-slate-400 italic">Không có trùng lặp</td></tr>`;
                document.getElementById('conflictCount').innerText = conflictResults.length;

                // Show Modal
                document.getElementById('importResultModal').classList.remove('hidden');
                document.getElementById('importResultModal').classList.add('flex');
                
                // 4. Save New People (if any)
                if (newPeople.length > 0) {
                    utils.toggleLoader(true, 'Đang lưu...');
                    let defaultOrgId = state.userOrgId; let defaultOrgName = state.userOrgId;
                    if(state.userRole === 'dispatcher') {
                        const sel = document.getElementById('importOrgSelect'); defaultOrgId = sel.value; defaultOrgName = sel.options[sel.selectedIndex].text;
                    } else {
                        const found = state.orgs.find(o => o.id === defaultOrgId); if(found) defaultOrgName = found.name;
                    }

                    const batch = db.batch();
                    const collectionRef = getPublicColl('dailyPlans').doc(state.currentDateKey).collection('people');
                    const planRef = getPublicColl('dailyPlans').doc(state.currentDateKey);
                    batch.set(planRef, { updatedAt: new Date(), status: state.isPlanLocked ? 'Locked' : 'Draft' }, { merge: true });

                    newPeople.forEach(p => {
                        const newRef = collectionRef.doc();
                const warnings = [];
                        if(!p.staffNo) warnings.push('Thiếu danh số');
                        if(p.idExpiryDate) {
                             const parts = p.idExpiryDate.split(/[\/\-\.]/);
                             if(parts.length === 3) {
                                 const exp = new Date(parseInt(parts[2].length===2?'20'+parts[2]:parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
                                 const diff = Math.ceil((exp - new Date()) / (86400000));
                                 if(diff < 0) warnings.push('CCCD Hết hạn'); else if(diff < 30) warnings.push('CCCD sắp hết hạn');
                             }
                        }
                        
                        // CHANGE: Force defaultOrgName for non-dispatcher users even if Excel has value
                        const finalOrgName = defaultOrgName;
                        
                        batch.set(newRef, {
                            ...p, orgName: finalOrgName, dkDate: date || '', destination: destination || '', nvsxNo: utils.normalizeNvsxInput(nvsx || ''), taskDesc: utils.normalizeTaskText(task || ''),
                            orgId: defaultOrgId, warnings, createdAt: new Date(), importedBy: state.currentUser.email
                        });
                    });

                    try {
                        await batch.commit();
                        audit.log('IMPORT', { count: newPeople.length, destination: destination || '', nvsxNo: nvsx || '', reason: 'import_excel' });
                        masterDataManager.saveToMaster(newPeople); 
                        utils.showToast(`Đã import xong`, 'success');
                        importModal.close();
                    } catch(e) { utils.showToast('Lỗi lưu: ' + e.message, 'error'); } finally { utils.toggleLoader(false); }
                } else {
                     importModal.close(); // Close import modal if nothing to save
                }
            }
        };

        const editModal = {
            currentPersonId: null,
            // NEW: Function to render destination suggestions
            renderDestinationSuggestions: () => {
                const uniqueDestinations = [...new Set(state.currentPlanPeople.map(p => p.destination).filter(Boolean))];
                const container = document.getElementById('destSuggestions');
                if (uniqueDestinations.length === 0) {
                    container.innerHTML = '<span class="text-slate-400 italic">Chưa có giàn nào hôm nay</span>';
                    return;
                }
                container.innerHTML = uniqueDestinations.map(dest => 
                    `<span onclick="document.getElementById('edit_destination').value = '${dest}'" class="cursor-pointer bg-blue-100 hover:bg-blue-200 text-blue-800 px-2 py-1 rounded border border-blue-200 transition">${dest}</span>`
                ).join('');
            },
            open: async (personId = null) => {
                utils.toggleLoader(true, "Đang xử lý...");
                editModal.currentPersonId = personId;
                const titleEl = document.getElementById('editModalTitle');
                
                ['edit_staffNo','edit_fullName','edit_title','edit_orgName','edit_dob','edit_pob','edit_phone','edit_idNo',
                 'edit_idIssueDate','edit_idExpiryDate','edit_weightKg','edit_luggageKg','edit_cargoKg','edit_rowNote',
                 'edit_destination','edit_nvsxNo','edit_taskDesc'].forEach(id => {
                     // Handle resetting input or select
                     const el = document.getElementById(id);
                     if(el) el.value = '';
                 });
                
                // Clear suggestions initially
                document.getElementById('destSuggestions').innerHTML = '';

                // DYNAMIC ORG FIELD CONSTRUCTION
                const orgContainer = document.getElementById('orgInputContainer');
                orgContainer.innerHTML = ''; // Clear previous

                if (state.userRole === 'dispatcher') {
                    // Admin: Input + Datalist
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.id = 'edit_orgName';
                    input.className = 'w-full border rounded p-2 text-sm';
                    input.setAttribute('list', 'allOrgsDatalist');
                    input.placeholder = 'Chọn hoặc nhập tên đơn vị...';
                    orgContainer.appendChild(input);
                } else {
                    // User: Select Only
                    const select = document.createElement('select');
                    select.id = 'edit_orgName';
                    select.className = 'w-full border rounded p-2 text-sm bg-slate-50';
                    
                    // Add options
                    state.orgs.forEach(org => {
                        const opt = document.createElement('option');
                        opt.value = org.name;
                        opt.text = org.name;
                        select.appendChild(opt);
                    });
                    
                    // Pre-select user's org
                    const userOrgName = state.orgs.find(o => o.id === state.userOrgId)?.name || state.userOrgId;
                    select.value = userOrgName;
                    orgContainer.appendChild(select);
                }

                try {
                    if (personId) {
                        titleEl.innerText = "Sửa thông tin";
                        // UPDATED PATH: Use public/data/dailyPlans
                        const doc = await getPublicColl('dailyPlans').doc(state.currentDateKey).collection('people').doc(personId).get();
                        if(!doc.exists) throw new Error("Không tìm thấy nhân sự");
                        const p = doc.data();
                        
                        document.getElementById('edit_staffNo').value = p.staffNo || '';
                        document.getElementById('edit_fullName').value = p.fullName || '';
                        document.getElementById('edit_title').value = p.title || '';
                        
                        // Set Org Name
                        const orgField = document.getElementById('edit_orgName');
                        if (orgField) orgField.value = p.orgName || '';
                        
                        document.getElementById('edit_dob').value = p.dob || '';
                        document.getElementById('edit_pob').value = p.pob || '';
                        document.getElementById('edit_phone').value = p.phone || '';
                        document.getElementById('edit_idNo').value = p.idNo || '';
                        document.getElementById('edit_idIssueDate').value = p.idIssueDate || '';
                        document.getElementById('edit_idExpiryDate').value = p.idExpiryDate || '';
                        document.getElementById('edit_weightKg').value = p.weightKg || '';
                        document.getElementById('edit_luggageKg').value = p.luggageKg || '';
                        document.getElementById('edit_cargoKg').value = p.cargoKg || '';
                        document.getElementById('edit_rowNote').value = p.rowNote || '';
                        document.getElementById('edit_destination').value = p.destination || '';
                        document.getElementById('edit_nvsxNo').value = p.nvsxNo || '';
                        document.getElementById('edit_taskDesc').value = p.taskDesc || '';
                        document.getElementById('advEdit').classList.add('hidden');
                    } else {
                        titleEl.innerText = "Thêm nhân sự mới";
                        // Default show advanced in add mode so suggestions are visible
                        document.getElementById('advEdit').classList.remove('hidden');
                        editModal.renderDestinationSuggestions(); // Load suggestions immediately
                    }
                    document.getElementById('editPersonModal').classList.remove('hidden'); document.getElementById('editPersonModal').classList.add('flex');
                } catch(e) { utils.showToast(e.message, "error"); } finally { utils.toggleLoader(false); }
            },
            close: () => { document.getElementById('editPersonModal').classList.add('hidden'); document.getElementById('editPersonModal').classList.remove('flex'); },
            save: async () => {
                utils.toggleLoader(true, "Đang lưu...");
                
                // Get Org Name (Input for Admin, Select for User)
                let finalOrgName = document.getElementById('edit_orgName').value;
                
                // Fallback for safety (though select should always have value)
                if (!finalOrgName && state.userRole !== 'dispatcher') {
                     const org = state.orgs.find(o => o.id === state.userOrgId);
                     finalOrgName = org ? org.name : state.userOrgId;
                }
                
                const updatedData = {
                    staffNo: document.getElementById('edit_staffNo').value, fullName: document.getElementById('edit_fullName').value,
                    title: document.getElementById('edit_title').value, dob: document.getElementById('edit_dob').value,
                    pob: document.getElementById('edit_pob').value, phone: document.getElementById('edit_phone').value,
                    idNo: document.getElementById('edit_idNo').value, idIssueDate: document.getElementById('edit_idIssueDate').value,
                    idExpiryDate: document.getElementById('edit_idExpiryDate').value,
                    weightKg: Number(document.getElementById('edit_weightKg').value), luggageKg: Number(document.getElementById('edit_luggageKg').value),
                    cargoKg: Number(document.getElementById('edit_cargoKg').value), rowNote: document.getElementById('edit_rowNote').value,
                    destination: document.getElementById('edit_destination').value, nvsxNo: document.getElementById('edit_nvsxNo').value,
                    taskDesc: document.getElementById('edit_taskDesc').value, 
                    orgName: finalOrgName,
                    updatedBy: state.currentUser.email, updatedAt: new Date()
                };

                const warnings = [];
                if(!updatedData.staffNo) warnings.push('Thiếu danh số');
                if(updatedData.idExpiryDate) {
                    const parts = updatedData.idExpiryDate.split('/');
                    if(parts.length === 3) {
                        const expDate = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
                        const diffDays = Math.ceil((expDate - new Date()) / (86400000));
                        if(diffDays <= 60) warnings.push('CCCD sắp/đã hết hạn'); updatedData.isIdExpiring = true;
                    }
                }
                updatedData.warnings = warnings;

                // Conflict Check
                if (!editModal.currentPersonId) { // Only check on create or if Staff ID changed (simplified to create for now)
                    const conflicts = conflictChecker.check([updatedData], state.currentPlanPeople);
                    if(conflicts.length > 0) {
                        utils.showToast(conflicts[0].msg, 'error'); // Use .msg property
                        utils.toggleLoader(false);
                        return;
                    }
                }

                try {
                    // UPDATED PATH: Use public/data/dailyPlans
                    if (editModal.currentPersonId) {
                        await getPublicColl('dailyPlans').doc(state.currentDateKey).collection('people').doc(editModal.currentPersonId).update(updatedData);
                        audit.log('UPDATE_PERSON', { staffNo: updatedData.staffNo, fullName: updatedData.fullName, destination: updatedData.destination });
                        utils.showToast("Đã cập nhật", "success");
                    } else {
                        updatedData.createdAt = new Date(); updatedData.importedBy = state.currentUser.email; updatedData.orgId = state.userRole === 'dispatcher' ? 'XNXL' : state.userOrgId;
                        await getPublicColl('dailyPlans').doc(state.currentDateKey).set({ updatedAt: new Date(), status: state.isPlanLocked ? 'Locked' : 'Draft' }, { merge: true });
                        await getPublicColl('dailyPlans').doc(state.currentDateKey).collection('people').add(updatedData);
                        audit.log('ADD_PERSON', { staffNo: updatedData.staffNo, fullName: updatedData.fullName, destination: updatedData.destination });
                        utils.showToast("Đã thêm mới", "success");
                    }
                    masterDataManager.saveToMaster([updatedData]); // Save to Master
                    editModal.close();
                } catch(e) { console.error(e); utils.showToast("Lỗi: " + e.message, "error"); } finally { utils.toggleLoader(false); }
            }
        };

        const exportManager = {
    selectedDates: [],
    pending: null,

    open: () => {
        if (exportManager.selectedDates.length === 0) {
            const current = document.getElementById('planDateInput').value;
            if (current) exportManager.addDateValue(current);
        }
        exportManager.renderDateList();
        document.getElementById('exportModal').classList.remove('hidden');
        document.getElementById('exportModal').classList.add('flex');
    },
    close: () => {
        document.getElementById('exportModal').classList.add('hidden');
        document.getElementById('exportModal').classList.remove('flex');
    },

    openPreview: () => {
        document.getElementById('exportPreviewModal').classList.remove('hidden');
        document.getElementById('exportPreviewModal').classList.add('flex');
    },
    closePreview: () => {
        document.getElementById('exportPreviewModal').classList.add('hidden');
        document.getElementById('exportPreviewModal').classList.remove('flex');
    },
    cancelPreview: () => {
        exportManager.closePreview();
        exportManager.open();
    },

    addDate: () => { exportManager.addDateValue(document.getElementById('exportDateInput').value); },
    addDateValue: (dateStr) => {
        if (dateStr && !exportManager.selectedDates.includes(dateStr)) {
            exportManager.selectedDates.push(dateStr);
            exportManager.selectedDates.sort();
            exportManager.renderDateList();
        }
    },
    removeDate: (dateStr) => {
        exportManager.selectedDates = exportManager.selectedDates.filter(d => d !== dateStr);
        exportManager.renderDateList();
    },
    renderDateList: () => {
        document.getElementById('exportDateList').innerHTML = exportManager.selectedDates.map(dateStr => {
            const [y, m, d] = dateStr.split('-');
            return `<li class="flex justify-between items-center bg-white p-2 rounded border border-slate-200 text-sm"><span class="font-bold text-slate-700">${d}/${m}/${y}</span><button onclick="exportManager.removeDate('${dateStr}')" class="text-red-500 hover:text-red-700 font-bold px-2">&times;</button></li>`;
        }).join('');
    },

    renderPreview: (peopleData, dateDisplayString) => {
        const esc = (s) => String(s ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');

        const people = [...peopleData];
        people.sort((a, b) => (a.destination || '').localeCompare(b.destination || '') || (a.nvsxNo || '').localeCompare(b.nvsxNo || '') || (a.orgName || '').localeCompare(b.orgName || ''));

        const grouped = people.reduce((acc, p) => {
            const k = p.destination || 'KHÁC';
            if (!acc[k]) acc[k] = [];
            acc[k].push(p);
            return acc;
        }, {});

        let html = `
            <div class="bg-slate-50 border border-slate-200 rounded-lg p-4">
                <div class="text-slate-700 font-bold mb-1"><i class="fas fa-calendar-day text-slate-400 mr-2"></i>Ngày khởi hành: <span class="text-blue-700">${esc(dateDisplayString)}</span></div>
                <div class="text-xs text-slate-500">Danh sách dưới đây đúng theo nội dung sẽ có trong file Word.</div>
            </div>
        `;

        Object.keys(grouped).forEach(dest => {
            const group = grouped[dest];
            const meta = utils.getGroupFinalMeta(group);
            const uniqueNVSX = meta.nvsxStr;
            const uniqueTasks = (meta.taskLines || []).join('\n');
            const isInconsistent = !!meta.inconsistent;

            html += `
                <div class="border border-slate-200 rounded-lg overflow-hidden">
                    <div class="bg-blue-50 px-4 py-3 flex justify-between items-center">
                        <div class="font-bold text-blue-800 text-base flex items-center">
                            <i class="fas fa-map-marker-alt mr-2 text-blue-500"></i> ${esc(dest)}
                        </div>
                        <div class="text-xs bg-white text-blue-600 px-2 py-1 rounded font-bold border border-blue-100">${group.length} người</div>
                    </div>
                    <div class="p-4 space-y-3">
                        <div class="text-sm text-slate-700"><span class="font-bold">Số NVSX:</span> ${esc(uniqueNVSX || '--')}</div>
                        <div class="text-sm text-slate-700">
                            <div class="font-bold mb-1">Nội dung công việc:</div>
                            <pre class="whitespace-pre-wrap text-xs bg-white p-3 rounded border border-slate-200 text-slate-700 font-mono">${esc(uniqueTasks || '--')}</pre>
                        </div>

                        <div class="text-sm text-slate-700 font-bold">Danh sách người đi:</div>
                        <div class="overflow-x-auto border border-slate-200 rounded">
                            <table class="min-w-full divide-y divide-slate-200 text-xs bg-white">
                                <thead class="bg-slate-50">
                                    <tr>
                                        <th class="px-3 py-2 text-center w-12">STT</th>
                                        <th class="px-3 py-2 text-left">Họ và tên</th>
                                        <th class="px-3 py-2 text-left">Chức danh</th>
                                        <th class="px-3 py-2 text-left">Đơn vị</th>
                                        <th class="px-3 py-2 text-center w-24">Danh số</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-100">
                                    ${group.map((p, i) => `
                                        <tr>
                                            <td class="px-3 py-2 text-center text-slate-500 font-mono">${i + 1}</td>
                                            <td class="px-3 py-2 font-semibold text-slate-700">${esc(p.fullName || '')}</td>
                                            <td class="px-3 py-2 text-slate-600">${esc(p.title || '')}</td>
                                            <td class="px-3 py-2 text-slate-600">${esc(importModal.getSelectedOrgName() || p.orgName || '')}</td>
                                            <td class="px-3 py-2 text-center text-slate-600 font-mono">${esc(p.staffNo || '')}</td>
                                        </tr>
                                    `).join('')}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            `;
        });

        document.getElementById('exportPreviewBody').innerHTML = html;
    },

    processExport: async () => {
        if (exportManager.selectedDates.length === 0) return utils.showToast("Vui lòng chọn ít nhất 1 ngày", "error");

        utils.toggleLoader(true, "Đang tổng hợp...");
        try {
            const promises = exportManager.selectedDates.map(async dateStr => {
                // UPDATED PATH: Use public/data/dailyPlans
                const snap = await getPublicColl('dailyPlans').doc(dateStr.replace(/-/g, '')).collection('people').get();
                return snap.docs.map(d => d.data());
            });
            const results = await Promise.all(promises);
            const allPeople = results.flat();
            if (allPeople.length === 0) { utils.showToast("Không có dữ liệu", "error"); utils.toggleLoader(false); return; }

            // Warn if group data is inconsistent (still allow export)
            try {
                const g = allPeople.reduce((acc, p) => { const k = (p.destination || 'KHÁC'); (acc[k] = acc[k] || []).push(p); return acc; }, {});
                const hasInconsistent = Object.keys(g).some(k => utils.getGroupFinalMeta(g[k]).inconsistent);
                if (hasInconsistent) {
                    utils.showToast("Cảnh báo: Dữ liệu không đồng nhất trong cùng 1 giàn. Nên bấm Sửa giàn/NVSX/CV để chốt lại trước khi Export (vẫn có thể Export).", "warning");
                }
            } catch(e) {}

            const dateDisplayString = exportManager.selectedDates
                .map(ds => { const [y, m, d] = ds.split('-'); return `${d}.${m}.${y}`; })
                .join(' + ');

            exportManager.pending = { peopleData: allPeople, dateDisplayString };

            // Show preview modal
            exportManager.close();
            exportManager.renderPreview(allPeople, dateDisplayString);
            exportManager.openPreview();
        } catch (e) {
            console.error(e);
            utils.showToast("Lỗi: " + e.message, "error");
        } finally {
            utils.toggleLoader(false);
        }
    },

    confirmExport: async () => {
        if (!exportManager.pending) return;

        utils.toggleLoader(true, "Đang tạo file...");
        try {
            await app.generateDocx(exportManager.pending.peopleData, exportManager.pending.dateDisplayString);
            audit.log('EXPORT_DOCX', { count: exportManager.pending.peopleData.length, reason: exportManager.pending.dateDisplayString });
            utils.showToast("Đã xuất file thành công", "success");
            exportManager.pending = null;
            exportManager.closePreview();
        } catch (e) {
            console.error(e);
            utils.showToast("Lỗi: " + e.message, "error");
        } finally {
            utils.toggleLoader(false);
        }
    }

,
    confirmPdf: async () => {
        if (!exportManager.pending) return;

        utils.toggleLoader(true, "Đang tạo PDF...");
        try {
            await app.generatePdf(exportManager.pending.peopleData, exportManager.pending.dateDisplayString);
            audit.log('EXPORT_PDF', { count: exportManager.pending.peopleData.length, reason: exportManager.pending.dateDisplayString });
            utils.showToast("Đã xuất PDF thành công", "success");
            exportManager.pending = null;
            exportManager.closePreview();
        } catch (e) {
            console.error(e);
            utils.showToast("Lỗi: " + e.message, "error");
        } finally {
            utils.toggleLoader(false);
        }
    }

};

        const app = {
            init: () => {
                auth.onAuthStateChanged(async (user) => {
                    if (user) {
                        utils.toggleLoader(true, 'Đang tải...');
                        const userDoc = await db.collection('artifacts').doc(appId).collection('users').doc(user.uid).get();
                        await db.collection('artifacts').doc(appId).collection('users').doc(user.uid).set({ lastLogin: new Date() }, { merge: true });
                        
                        if (!userDoc.exists) {
                            const role = user.email.includes('admin') ? 'dispatcher' : 'workshop'; const orgId = role === 'dispatcher' ? 'XNXL' : 'XuongBo';
                            await db.collection('artifacts').doc(appId).collection('users').doc(user.uid).set({ email: user.email, role, orgId, displayName: user.displayName || user.email, createdAt: new Date() });
                            state.userRole = role; state.userOrgId = orgId;
                        } else {
                            const d = userDoc.data(); state.userRole = d.role; state.userOrgId = d.orgId;
                        }
                        state.currentUser = user; app.setupUI(); app.displayLastLoginInfo();
                        document.getElementById('authScreen').classList.add('hidden'); document.getElementById('appShell').classList.remove('hidden'); document.getElementById('appShell').classList.add('flex');
                        
                        await app.loadOrgs(); if(state.userRole === 'dispatcher') userManager.loadUsers();
                        
                        const today = new Date(); const mm = String(today.getMonth()+1).padStart(2,'0'); const dd = String(today.getDate()).padStart(2,'0'); const yyyy = today.getFullYear();
                        document.getElementById('planDateInput').value = `${yyyy}-${mm}-${dd}`; document.getElementById('exportDateInput').value = `${yyyy}-${mm}-${dd}`;
                        state.currentDateKey = `${yyyy}${mm}${dd}`;
                        
                        app.loadPlan(); // Start listener
                        utils.toggleLoader(false);
                    } else {
                        document.getElementById('authScreen').classList.remove('hidden'); document.getElementById('appShell').classList.add('hidden'); document.getElementById('appShell').classList.remove('flex');
                    }
                });
                document.getElementById('loginForm').addEventListener('submit', async (e) => {
                    e.preventDefault();
                    let email = document.getElementById('email').value.trim(); if (!email.includes('@')) email += '@xnxl.com';
                    try { await auth.signInWithEmailAndPassword(email, document.getElementById('password').value); } catch (err) { utils.showToast(err.message, 'error'); }
                });
                document.getElementById('excelFile').addEventListener('change', importModal.handleFile);
                // Quick search filter (client-side)
                const qs = document.getElementById('quickSearch');
                const qsClear = document.getElementById('btnClearSearch');
                if (qs) {
                    qs.addEventListener('input', () => {
                        state.quickSearch = qs.value || '';
                        app.renderPlanTable(state.currentPlanPeople || []);
                    });
                }
                if (qsClear && qs) {
                    qsClear.addEventListener('click', () => {
                        qs.value = '';
                        state.quickSearch = '';
                        app.renderPlanTable(state.currentPlanPeople || []);
                    });
                }
                // Warning center click
                const warnBox = document.getElementById('warningStatBox');
                if (warnBox) {
                    warnBox.style.cursor = 'pointer';
                    warnBox.title = 'Bấm để xem chi tiết cảnh báo';
                    warnBox.addEventListener('click', () => {
                        const n = parseInt(document.getElementById('statWarnings')?.innerText || '0', 10);
                        if (n > 0) warningCenter.open();
                        else utils.showToast('Không có cảnh báo', 'success');
                    });
                }

                // Responsive (mobile card / desktop table): re-render when screen size changes
                if (!window.__planResizeBound) {
                    window.__planResizeBound = true;
                    let __rt;
                    window.addEventListener('resize', () => {
                        clearTimeout(__rt);
                        __rt = setTimeout(() => {
                            try {
                                app.renderPlanTable(state.currentPlanPeople || []);
                            } catch (e) { /* ignore */ }
                        }, 200);
                    });
                }

            },
            logout: async () => { if(state.unsubscribePlan) state.unsubscribePlan(); await auth.signOut(); },
            displayLastLoginInfo: async () => {
                const loginValEl = document.getElementById('lastLoginVal');
                if (state.userRole === 'dispatcher') {
                    try {
                        const snap = await db.collection('artifacts').doc(appId).collection('users').get();
                        let latestTime = 0, latestUser = '';
                        snap.forEach(doc => { const d = doc.data(); if (d.lastLogin && d.lastLogin.seconds > latestTime/1000) { latestTime = d.lastLogin.seconds*1000; latestUser = d.displayName || d.email; } });
                        if (latestTime > 0) loginValEl.innerText = `${latestUser} (${utils.formatDateTime(new Date(latestTime))})`;
                    } catch (e) { loginValEl.innerText = "Lỗi tải"; }
                } else { loginValEl.innerText = utils.formatDateTime(new Date()); }
            },
            setupUI: () => {
                document.getElementById('userBadge').innerText = state.userRole === 'dispatcher' ? 'ĐIỀU ĐỘ (ADMIN)' : `USER: ${state.userOrgId}`;
                const isDisp = state.userRole === 'dispatcher';
                document.getElementById('navSettings').classList.toggle('hidden', !isDisp);                 document.getElementById('navAudit').classList.toggle('hidden', !isDisp);
document.getElementById('btnLock').classList.toggle('hidden', !isDisp);
                document.getElementById('importOrgSelectWrapper').classList.toggle('hidden', !isDisp); document.getElementById('btnExport').classList.toggle('hidden', !isDisp);
                document.getElementById('btnReport').classList.toggle('hidden', !isDisp); // Toggle Report Button
            },
            loadOrgs: async () => {
                // UPDATED PATH: Use public/data/organizations
                const snap = await getPublicColl('organizations').get();
                if (snap.empty) { const seeds = ['Xưởng Bờ','Xưởng Sửa Chữa','Ban Chánh Hàn','Xưởng Biển','Ban Khảo Sát','XNXL']; const batch = db.batch(); seeds.forEach(n => batch.set(getPublicColl('organizations').doc(n), {name:n})); await batch.commit(); state.orgs = seeds.map(n => ({id:n,name:n})); }
                else state.orgs = snap.docs.map(d => ({id:d.id,...d.data()}));
                const ops = state.orgs.map(o => `<option value="${o.name}">${o.name}</option>`).join(''); // Use Name as value for datalist consistency
                // Update Datalist for Admin
                document.getElementById('allOrgsDatalist').innerHTML = ops;
                // Update Import Select for Admin
                document.getElementById('importOrgSelect').innerHTML = state.orgs.map(o => `<option value="${o.id}">${o.name}</option>`).join('');
                document.getElementById('u_orgId').innerHTML = state.orgs.map(o => `<option value="${o.id}">${o.name}</option>`).join('');
            },
            changeDate: () => {
                const dateVal = document.getElementById('planDateInput').value;
                state.currentDateKey = dateVal.replace(/-/g, '');
                app.loadPlan();
            },
            loadPlan: () => {
                if (state.unsubscribePlan) state.unsubscribePlan(); // Unsubscribe prev listener
                utils.toggleLoader(true);
                
                // 1. Get Status (Single fetch to avoid 2 listeners)
                // UPDATED PATH: Use public/data/dailyPlans
                getPublicColl('dailyPlans').doc(state.currentDateKey).get().then(doc => {
                    state.isPlanLocked = false;
                    if(doc.exists) {
                        state.isPlanLocked = doc.data().status === 'Locked';
                        document.getElementById('planStatus').innerText = state.isPlanLocked ? 'ĐÃ KHÓA' : 'DRAFT';
                        document.getElementById('planStatus').className = `px-3 py-1.5 rounded text-xs font-bold uppercase tracking-wider ${state.isPlanLocked ? 'bg-red-50 text-red-600 border border-red-100' : 'bg-green-50 text-green-700 border border-green-100'}`;
                    } else {
                        document.getElementById('planStatus').innerText = 'MỚI';
                        document.getElementById('planStatus').className = 'px-3 py-1.5 rounded text-xs font-bold bg-slate-50 text-slate-500 border border-slate-200';
                    }
                    // Update Button States
                    const isDisp = state.userRole === 'dispatcher';
                    document.getElementById('btnImport').disabled = state.isPlanLocked && !isDisp;
                    document.getElementById('btnImport').classList.toggle('opacity-50', state.isPlanLocked && !isDisp);
                    document.getElementById('btnAdd').disabled = state.isPlanLocked && !isDisp;
                    document.getElementById('btnAdd').classList.toggle('opacity-50', state.isPlanLocked && !isDisp);
                    document.getElementById('btnLock').innerHTML = state.isPlanLocked ? '<i class="fas fa-unlock mr-1"></i> Mở' : '<i class="fas fa-lock mr-1"></i> Khóa';
                    document.getElementById('btnDeleteAll').classList.toggle('hidden', state.isPlanLocked && !isDisp);
                });

                // 2. Real-time Listen People
                // UPDATED PATH: Use public/data/dailyPlans
                let q = getPublicColl('dailyPlans').doc(state.currentDateKey).collection('people');
                if (state.userRole !== 'dispatcher') q = q.where('orgId', '==', state.userOrgId);

                state.unsubscribePlan = q.onSnapshot(snap => {
                    const people = snap.docs.map(d => ({...d.data(), id: d.id}));
                    state.currentPlanPeople = people;
                    app.renderPlanTable(people);
                    utils.toggleLoader(false);
                }, err => { console.error(err); utils.toggleLoader(false); });
            },
            renderPlanTable: (people) => {
                const container = document.getElementById('planContent');
                const allPeople = Array.isArray(people) ? people : [];
                const q = (state.quickSearch || '').toLowerCase().trim();
                const viewPeople = q ? allPeople.filter(p => {
                    const hay = [
                        p.staffNo, p.fullName, p.title, p.orgName, p.destination, p.nvsxNo, p.taskDesc
                    ].join(' ').toLowerCase();
                    return hay.includes(q);
                }) : allPeople;
                if (allPeople.length === 0) {
                    container.innerHTML = `<div class="flex flex-col items-center justify-center h-64 text-slate-400 bg-white"><div class="bg-slate-50 p-6 rounded-full mb-4"><i class="fas fa-clipboard-list text-4xl text-slate-300"></i></div><p class="text-sm font-medium">Chưa có dữ liệu.</p></div>`;
                    app.updateStats(0,0,0,0); return;
                }
                const offshore = people.filter(p => !p.isReturn).length;
                const warnings = allPeople.filter(p => p.warnings && p.warnings.length > 0).length;
                app.updateStats(allPeople.length, allPeople.length, 0, warnings); // Assuming all go offshore for now
                if (viewPeople.length === 0) {
                    container.innerHTML = `<div class="flex flex-col items-center justify-center h-64 text-slate-400 bg-white"><div class="bg-slate-50 p-6 rounded-full mb-4"><i class="fas fa-search text-4xl text-slate-300"></i></div><p class="text-sm font-medium">Không có kết quả tìm kiếm</p><p class="text-xs mt-1">Hãy thử từ khóa khác (tên, danh số, giàn, đơn vị...)</p></div>`;
                    return;
                }

                const grouped = viewPeople.reduce((acc, p) => { const k = utils.destinationDisplay(p.destination); if(!acc[k]) acc[k]=[]; acc[k].push(p); return acc; }, {});
                let html = '';
                                const isMobile = !!(window.matchMedia && window.matchMedia('(max-width: 767px)').matches);

                Object.keys(grouped).sort().forEach(dest => {
                    const group = grouped[dest];
                    const meta = utils.getGroupFinalMeta(group);
                    const allNvsx = meta.nvsxStr;
                    const taskLines = meta.taskLines;
                    const isInconsistent = !!meta.inconsistent;

                    const escHtml = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
                    const taskDisplayHtml = taskLines.length ? taskLines.map(l => escHtml(l)).join('<br/>') : '';

                    const gid = 'task_' + btoa(unescape(encodeURIComponent(dest))).replace(/=+/g,'').replace(/[^a-zA-Z0-9]/g,'');

                    const safeDest = dest.replace(/'/g, "\'");
                    const editBtn = state.userRole === 'dispatcher'
                        ? `<button onclick="app.editDestination('${safeDest}')" class="ml-2 text-slate-400 hover:text-blue-600 transition"><i class="fas fa-edit"></i></button>`
                        : '';

                    html += `
                        <div class="bg-blue-50/50 px-4 py-3 border-b border-blue-100 sticky top-0 z-10 backdrop-blur-sm shadow-sm">
                            <div class="flex justify-between items-start">
                                <div class="w-4/5">
                                    <div class="font-bold text-blue-800 text-base flex items-center">
                                        <i class="fas fa-map-marker-alt mr-2 text-blue-500"></i> ${dest} ${editBtn}
                                    </div>
                                    <div class="mt-1 text-xs text-slate-600 space-y-1">
                                        ${isInconsistent ? `<div class="text-xs text-amber-800 bg-amber-50 border border-amber-200 rounded-md p-2">Dữ liệu không đồng nhất, hãy bấm <b>Sửa giàn/NVSX/CV</b> để chốt lại (vẫn có thể Export).</div>` : ``}
                                        <div><span class="font-bold">NVSX:</span> <span class="break-words">${allNvsx || '--'}</span></div>
                                        <div class="flex items-start">
                                            <span class="font-bold mr-1 whitespace-nowrap">Công việc:</span>
                                            <div class="flex-1">
                                                <div id="${gid}" class="js-task-block text-slate-700 leading-relaxed" data-lines="${taskLines.length}">${taskDisplayHtml || '--'}</div>
                                                <button type="button" class="js-task-toggle hidden mt-1 text-[11px] text-blue-700 font-bold hover:underline" data-target="${gid}">Mở rộng</button>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <span class="text-xs bg-white text-blue-600 px-2 py-1 rounded font-bold border border-blue-100">${group.length} người</span>
                            </div>
                        </div>
                    `;

                    const sortedGroup = [...group].sort((a,b) => (a.nvsxNo||'').localeCompare(b.nvsxNo||'') || (a.staffNo||'').localeCompare(b.staffNo||''));

                    if (isMobile) {
                        html += `<div class="bg-white divide-y divide-slate-100">`;
                        sortedGroup.forEach((p, idx) => {
                            const warn = (p.warnings && p.warnings.length > 0);
                            const warnIcon = warn ? `<i class="fas fa-exclamation-triangle text-red-500 ml-1" title="${(p.warnings||[]).join(', ')}"></i>` : '';
                            const canEdit = state.userRole === 'dispatcher' || (!state.isPlanLocked && p.orgId === state.userOrgId);

                            let updatedInfo = '';
                            if (p.updatedAt || p.createdAt) {
                                const t = p.updatedAt ? p.updatedAt.seconds*1000 : p.createdAt.seconds*1000;
                                const u = (p.updatedBy || p.importedBy || '').split('@')[0];
                                updatedInfo = `<div class="text-[10px] text-slate-400 font-mono mt-1">${u} ${utils.formatTimeShort(new Date(t))}</div>`;
                            }

                            const actionBtns = canEdit
                                ? `<div class="flex items-center gap-3"><button onclick="editModal.open('${p.id}')" class="text-blue-600 hover:text-blue-800"><i class="fas fa-pen"></i></button><button onclick="app.deletePerson('${p.id}')" class="text-red-500 hover:text-red-700"><i class="fas fa-trash-alt"></i></button></div>`
                                : `<div class="text-slate-300"><i class="fas fa-lock"></i></div>`;

                            html += `
                                <div class="p-3 ${warn ? 'bg-red-50/40' : ''}">
                                    <div class="flex items-start justify-between gap-3">
                                        <div class="min-w-0">
                                            <div class="flex items-center gap-2">
                                                <div class="text-xs text-slate-400 font-mono">${idx+1}</div>
                                                <div class="font-bold text-slate-800 truncate">${escHtml(p.fullName || '')}</div>
                                            </div>

                                            <div class="mt-0.5 text-xs text-slate-600 flex flex-wrap gap-x-3 gap-y-1">
                                                <div><span class="font-semibold">Danh số:</span> <span class="font-mono">${escHtml(p.staffNo || '')}</span></div>
                                                <div><span class="font-semibold">Chức danh:</span> ${escHtml(p.title || '')}</div>
                                            </div>

                                            <div class="mt-0.5 text-xs text-slate-600 flex flex-wrap gap-x-3 gap-y-1">
                                                <div><span class="font-semibold">Đơn vị:</span> ${escHtml(p.orgName || '-')}</div>
                                                <div><span class="font-semibold">SĐT:</span> <span class="font-mono">${escHtml(p.phone || '-')}</span></div>
                                            </div>

                                            <div class="mt-0.5 text-[11px] text-slate-500 flex flex-wrap gap-x-3 gap-y-1">
                                                <div><span class="font-semibold">CCCD:</span> <span class="font-mono">${escHtml(p.idNo || '-')}</span>${warnIcon}</div>
                                                <div><span class="font-semibold">Hết hạn:</span> <span class="${p.isIdExpiring?'text-red-600 font-bold':'text-slate-500'}">${escHtml(p.idExpiryDate || '')}</span></div>
                                            </div>

                                            ${p.rowNote ? `<div class="mt-1 text-[11px] italic text-slate-500 break-words"><span class="font-semibold not-italic">Ghi chú:</span> ${escHtml(p.rowNote)}</div>` : ``}
                                            ${updatedInfo}
                                        </div>

                                        <div class="flex flex-col items-end gap-2 flex-none">
                                            ${actionBtns}
                                            <div class="text-[11px] text-slate-500 text-right">
                                                <div><span class="font-semibold">KG:</span> ${escHtml(p.weightKg||0)}</div>
                                                <div><span class="font-semibold">HL:</span> ${escHtml(p.luggageKg||0)} | <span class="font-semibold">HH:</span> ${escHtml(p.cargoKg||0)}</div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            `;
                        });
                        html += `</div>`;
                    } else {
                        html += `
                            <div class="overflow-x-auto">
                                <table class="min-w-full divide-y divide-slate-100 text-xs">
                                    <thead>
                                        <tr>
                                            <th class="px-3 py-3 text-center w-10">STT</th>
                                            <th class="px-3 py-3 text-left w-20">Danh số</th>
                                            <th class="px-3 py-3 text-left w-40">Họ tên</th>
                                            <th class="px-3 py-3 text-left w-40">Chức danh / Đơn vị</th>
                                            <th class="px-3 py-3 text-left w-32">Ngày sinh / Nơi sinh</th>
                                            <th class="px-3 py-3 text-left w-24">SĐT</th>
                                            <th class="px-3 py-3 text-left w-32">CCCD / Hết hạn</th>
                                            <th class="px-3 py-3 text-center w-16">Cân nặng</th>
                                            <th class="px-3 py-3 text-center w-16">Hành lý</th>
                                            <th class="px-3 py-3 text-center w-16">Hàng hóa</th>
                                            <th class="px-3 py-3 text-left w-32">Ghi chú</th>
                                            <th class="px-3 py-3 text-center sticky right-0 bg-slate-50 w-20 shadow-l-sm">Sửa</th>
                                        </tr>
                                    </thead>
                                    <tbody class="bg-white divide-y divide-slate-50">
                        `;

                        sortedGroup.forEach((p, idx) => {
                            const warnClass = (p.warnings&&p.warnings.length>0) ? 'bg-red-50/50' : '';
                            const warnIcon = (p.warnings&&p.warnings.length>0) ? `<i class="fas fa-exclamation-triangle text-red-500 ml-1" title="${p.warnings.join(', ')}"></i>` : '';
                            const canEdit = state.userRole === 'dispatcher' || (!state.isPlanLocked && p.orgId === state.userOrgId);
                            let updatedInfo = '';
                            if(p.updatedAt || p.createdAt) {
                                const t = p.updatedAt ? p.updatedAt.seconds*1000 : p.createdAt.seconds*1000;
                                const u = (p.updatedBy || p.importedBy || '').split('@')[0];
                                updatedInfo = `<span class="text-[9px] text-slate-400 font-mono block mt-1 leading-tight">${u} ${utils.formatTimeShort(new Date(t))}</span>`;
                            }
                            const actionBtns = canEdit ? `<div class="flex flex-col items-center justify-center"><div class="flex space-x-2 mb-1"><button onclick="editModal.open('${p.id}')" class="text-blue-500 hover:text-blue-700 transition"><i class="fas fa-pen"></i></button><button onclick="app.deletePerson('${p.id}')" class="text-red-400 hover:text-red-600 transition"><i class="fas fa-trash-alt"></i></button></div>${updatedInfo}</div>` : `<div class="flex flex-col items-center"><span class="text-slate-300"><i class="fas fa-lock"></i></span>${updatedInfo}</div>`;
                            html += `<tr class="${warnClass} hover:bg-blue-50/30 transition duration-75 group"><td class="px-3 py-2 text-slate-400 text-center font-mono">${idx+1}</td><td class="px-3 py-2 text-slate-600 font-mono">${p.staffNo}</td><td class="px-3 py-2 font-semibold text-slate-700">${p.fullName}</td><td class="px-3 py-2 text-slate-600"><div class="font-medium">${p.title||''}</div><div class="text-[10px] text-slate-400 uppercase mt-0.5 font-bold tracking-wide">${p.orgName||'-'}</div></td><td class="px-3 py-2 text-slate-600"><div class="font-mono text-xs">${p.dob||'-'}</div><div class="text-[10px] text-slate-400 mt-0.5">${p.pob||''}</div></td><td class="px-3 py-2 text-slate-600 font-mono">${p.phone||'-'}</td><td class="px-3 py-2 text-slate-600"><div class="font-mono text-xs">${p.idNo||'-'}</div><div class="text-[10px] ${p.isIdExpiring?'text-red-600 font-bold':'text-slate-400'} mt-0.5">${p.idExpiryDate||''}${warnIcon}</div></td><td class="px-3 py-2 text-slate-600 text-center">${p.weightKg||0}</td><td class="px-3 py-2 text-slate-600 text-center">${p.luggageKg||0}</td><td class="px-3 py-2 text-slate-600 text-center">${p.cargoKg||0}</td><td class="px-3 py-2 text-slate-500 italic text-[11px] break-words max-w-[150px]">${p.rowNote||''}</td><td class="px-3 py-2 text-center sticky right-0 bg-white group-hover:bg-blue-50/30 shadow-l-sm border-l border-slate-50">${actionBtns}</td></tr>`;
                        });

                        html += `
                                    </tbody>
                                </table>
                            </div>
                        `;
                    }
                });
                container.innerHTML = html;
                setTimeout(() => { try { app.initTaskToggles(); } catch(e){} }, 0);
            },
            initTaskToggles: () => {
                try {
                    const blocks = document.querySelectorAll('.js-task-block');
                    blocks.forEach(el => {
                        const lines = parseInt(el.getAttribute('data-lines') || '0', 10);
                        const targetId = el.id;
                        const btn = el.parentElement?.querySelector('.js-task-toggle');
                        if (!btn || !targetId) return;

                        // Only enable toggle when more than 10 lines
                        if (lines > 10) {
                            el.classList.add('task-collapsed');
                            btn.classList.remove('hidden');
                            btn.textContent = 'Mở rộng';
                            btn.onclick = () => app.toggleTask(targetId);
                        } else {
                            el.classList.remove('task-collapsed');
                            btn.classList.add('hidden');
                        }
                    });
                } catch (e) {
                    console.warn('initTaskToggles failed', e);
                }
            },
            toggleTask: (targetId) => {
                const el = document.getElementById(targetId);
                if (!el) return;
                const btn = el.parentElement?.querySelector('.js-task-toggle');
                const collapsed = el.classList.contains('task-collapsed');
                if (collapsed) {
                    el.classList.remove('task-collapsed');
                    if (btn) btn.textContent = 'Thu gọn';
                } else {
                    el.classList.add('task-collapsed');
                    if (btn) btn.textContent = 'Mở rộng';
                }
            },
            editDestination: (oldName) => { 
    // Kept name for backwards-compat with existing UI button
    app.openGroupEditModal(oldName);
},

openGroupEditModal: (oldName) => {
    if (state.userRole !== 'dispatcher') return;

    const oldKey = utils.normalizeDestinationKey(oldName);
    const oldDisplay = utils.destinationDisplay(oldName);

    const targets = state.currentPlanPeople.filter(p => utils.normalizeDestinationKey(p.destination) === oldKey);
    if (targets.length === 0) return utils.showToast("Không tìm thấy nhân sự nào", "error");

    // Prefill with the first row (thường sau import sẽ giống nhau)
    document.getElementById('groupEditOldName').value = oldDisplay;
    document.getElementById('groupEditOldKey').value = oldKey;
    document.getElementById('groupEditDest').value = oldKey;
    document.getElementById('groupEditNvsx').value = (targets[0].nvsxNo || '');
    document.getElementById('groupEditTask').value = (targets[0].taskDesc || '');

    document.getElementById('groupEditModal').classList.remove('hidden');
    document.getElementById('groupEditModal').classList.add('flex');
},

closeGroupEditModal: () => {
    document.getElementById('groupEditModal').classList.add('hidden');
    document.getElementById('groupEditModal').classList.remove('flex');
},

saveGroupEditModal: async () => {
    const oldDisplay = document.getElementById('groupEditOldName').value || '';
    const oldKey = (document.getElementById('groupEditOldKey')?.value ?? '').toString();
    if (oldKey === undefined) return;

    const newDestInput = (document.getElementById('groupEditDest').value || '').trim();
    const newDest = newDestInput || oldKey;
    const newNvsx = (document.getElementById('groupEditNvsx').value || '').trim();
    const newTask = (document.getElementById('groupEditTask').value || '').trim();

    const targets = state.currentPlanPeople.filter(p => utils.normalizeDestinationKey(p.destination) === oldKey);
    if (targets.length === 0) return utils.showToast("Không tìm thấy nhân sự nào", "error");

    utils.toggleLoader(true, "Đang cập nhật...");
    try {
        const batch = db.batch();
        // UPDATED PATH: Use public/data/dailyPlans
        const col = getPublicColl('dailyPlans').doc(state.currentDateKey).collection('people');
        targets.forEach(p => batch.update(col.doc(p.id), {
            destination: newDest,
            nvsxNo: utils.normalizeNvsxInput(newNvsx),
            taskDesc: utils.normalizeTaskText(newTask),
            updatedBy: state.currentUser.email,
            updatedAt: new Date()
        }));
        await batch.commit();
        audit.log('GROUP_EDIT', { count: targets.length, destination: newDest, nvsxNo: newNvsx, reason: `from:${oldDisplay}` });
        utils.showToast(`Đã cập nhật giàn/NVSX/công việc cho ${targets.length} nhân sự`, 'success');
        app.closeGroupEditModal();
    } catch (e) {
        utils.showToast("Lỗi: " + e.message, "error");
    } finally {
        utils.toggleLoader(false);
    }
},
            deletePerson: async (id) => { 
                if(!confirm("Xác nhận xóa nhân sự này?")) return;
                utils.toggleLoader(true, "Đang xóa...");
                try {
                    await getPublicColl('dailyPlans').doc(state.currentDateKey).collection('people').doc(id).delete();
                    audit.log('DELETE_PERSON', { reason: 'delete_single', id });
                    utils.showToast("Đã xóa thành công", "success");
                } catch(e) {
                    utils.showToast("Lỗi: "+e.message, "error");
                } finally {
                    utils.toggleLoader(false);
                }
            },
            deleteAllPeople: async () => {
                if(state.currentPlanPeople.length === 0 || !confirm("CẢNH BÁO: Xóa TOÀN BỘ danh sách?") || !confirm("Chắc chắn xóa?")) return;
                utils.toggleLoader(true, 'Đang xóa...');
                try {
                    const batch = db.batch(); 
                    // UPDATED PATH: Use public/data/dailyPlans
                    const col = getPublicColl('dailyPlans').doc(state.currentDateKey).collection('people');
                    let targets = state.currentPlanPeople; if(state.userRole !== 'dispatcher') targets = targets.filter(p => p.orgId === state.userOrgId);
                    targets.forEach(p => batch.delete(col.doc(p.id))); await batch.commit();
                    audit.log('DELETE_ALL', { count: targets.length, reason: 'delete_all_people' });
                    utils.showToast(`Đã xóa ${targets.length} nhân sự`, "success");
                } catch(e) { utils.showToast("Lỗi: " + e.message, "error"); } finally { utils.toggleLoader(false); }
            },
            updateStats: (t, o, r, w) => {
                document.getElementById('statTotal').innerText = t; document.getElementById('statOffshore').innerText = o;
                document.getElementById('statReturn').innerText = r; document.getElementById('statWarnings').innerText = w;
                document.getElementById('warningStatBox').classList.toggle('opacity-30', w === 0);
            },
            toggleLockPlan: async () => {
                if(state.userRole !== 'dispatcher') return;
                const s = state.isPlanLocked ? 'Draft' : 'Locked';
                if (s === 'Locked') {
                    const w = (state.currentPlanPeople || []).filter(p => Array.isArray(p.warnings) && p.warnings.length > 0).length;
                    if (w > 0 && !confirm(`Còn ${w} cảnh báo (thiếu danh số/CCCD hết hạn...). Vẫn KHÓA ngày?`)) return;
                }
                // UPDATED PATH: Use public/data/dailyPlans
                await getPublicColl('dailyPlans').doc(state.currentDateKey).set({ status: s, lockedAt: new Date(), lockedBy: state.currentUser.email }, { merge: true });
                audit.log('LOCK_TOGGLE', { status: s, count: (state.currentPlanPeople || []).length });
                utils.showToast(`Đã ${s === 'Locked' ? 'KHÓA' : 'MỞ'} ngày đăng ký`, 'success');
                app.loadPlan(); // Manually reload status
            },
            showSettings: () => { document.getElementById('dashboardView').classList.add('hidden'); document.getElementById('settingsView').classList.remove('hidden'); app.renderOrgList(); app.loadSignatureConfig(); },
            showDashboard: () => { document.getElementById('settingsView').classList.add('hidden'); document.getElementById('dashboardView').classList.remove('hidden'); },
            renderOrgList: () => { document.getElementById('orgList').innerHTML = state.orgs.map(o => `<li class="flex justify-between items-center bg-slate-50 p-2 rounded border border-slate-100 text-xs"><span>${o.name}</span><button onclick="app.deleteOrg('${o.id}')" class="text-red-500 hover:text-red-700"><i class="fas fa-trash"></i></button></li>`).join(''); },
            addOrg: async () => { const n = document.getElementById('newOrgName').value.trim(); if(n) { await getPublicColl('organizations').doc(n).set({name:n, createdAt: new Date()}); document.getElementById('newOrgName').value=''; app.loadOrgs(); app.renderOrgList(); } },
            deleteOrg: async (id) => { if(confirm('Xóa đơn vị này?')) { await getPublicColl('organizations').doc(id).delete(); app.loadOrgs(); app.renderOrgList(); } },
            loadSignatureConfig: async () => { 
                // UPDATED PATH: Use public/data/settings
                const d = await getPublicColl('settings').doc('signature').get(); 
                if(d.exists) { const v = d.data(); document.getElementById('sigApprover').value = v.approver||''; document.getElementById('sigApproverName').value = v.approverName||''; } 
            },
            saveSettings: async () => { 
                // UPDATED PATH: Use public/data/settings
                await getPublicColl('settings').doc('signature').set({ approver: document.getElementById('sigApprover').value, approverName: document.getElementById('sigApproverName').value }); 
                utils.showToast('Đã lưu cấu hình'); 
            },
            
            generateDocx: async (peopleData, dateDisplayString) => {
                if (typeof docx === 'undefined' || typeof saveAs === 'undefined') {
                    throw new Error('Không tải được thư viện tạo Word (docx hoặc FileSaver). Hãy kiểm tra kết nối internet/CDN.');
                }

                let sig = { approver: 'TRƯỞNG BAN TTĐĐSX', approverName: '' };
                const sigDoc = await getPublicColl('settings').doc('signature').get();
                if(sigDoc.exists) sig = sigDoc.data();

                peopleData.sort((a, b) => (a.destination||'').localeCompare(b.destination||'') || (a.nvsxNo||'').localeCompare(b.nvsxNo||'') || (a.orgName||'').localeCompare(b.orgName||''));
                const grouped = peopleData.reduce((acc, p) => { const k = p.destination || 'KHÁC'; if(!acc[k]) acc[k]=[]; acc[k].push(p); return acc; }, {});

                const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, TextRun, AlignmentType, BorderStyle, HeightRule } = docx;
                const today = new Date(); const dd = String(today.getDate()).padStart(2,'0'); const mm = String(today.getMonth()+1).padStart(2,'0'); const yyyy = today.getFullYear();
                const FONT = "Times New Roman"; const SZ_N = 24; const SZ_S = 22;

                const children = [
                    new Table({
                        rows: [new TableRow({ children: [
                            new TableCell({ children: [
                                new Paragraph({ children: [new TextRun({ text: "LIÊN DOANH VIỆT – NGA", bold: true, size: SZ_S, font: FONT })], alignment: AlignmentType.CENTER }),
                                new Paragraph({ children: [new TextRun({ text: "VIETSOVPETRO", bold: true, size: SZ_S, font: FONT })], alignment: AlignmentType.CENTER }),
                                new Paragraph({ children: [new TextRun({ text: "XNXL", bold: true, size: SZ_S, font: FONT })], alignment: AlignmentType.CENTER }),
                                new Paragraph({ text: "" }),
                                new Paragraph({ children: [new TextRun({ text: "Số: ...../.....-CV-XL", size: SZ_S, font: FONT })], alignment: AlignmentType.CENTER })
                            ], borders: { top: {style: BorderStyle.NONE, size: 0, color: "auto"}, bottom: {style: BorderStyle.NONE, size: 0, color: "auto"}, left: {style: BorderStyle.NONE, size: 0, color: "auto"}, right: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideVertical: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideHorizontal: {style: BorderStyle.NONE, size: 0, color: "auto"} }, width: { size: 40, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [
                                new Paragraph({ children: [new TextRun({ text: "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold: true, size: SZ_S, font: FONT })], alignment: AlignmentType.CENTER }),
                                new Paragraph({ children: [new TextRun({ text: "Độc lập – Tự do – Hạnh phúc", bold: true, underline: { type: "single" }, size: SZ_S, font: FONT })], alignment: AlignmentType.CENTER }),
                                new Paragraph({ text: "" }),
                                new Paragraph({ children: [new TextRun({ text: `TPHCM, ngày ${dd} tháng ${mm} năm ${yyyy}`, size: SZ_S, italics: true, font: FONT })], alignment: AlignmentType.CENTER })
                            ], borders: { top: {style: BorderStyle.NONE, size: 0, color: "auto"}, bottom: {style: BorderStyle.NONE, size: 0, color: "auto"}, left: {style: BorderStyle.NONE, size: 0, color: "auto"}, right: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideVertical: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideHorizontal: {style: BorderStyle.NONE, size: 0, color: "auto"} }, width: { size: 60, type: WidthType.PERCENTAGE } })
                        ] })], width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: { top: {style: BorderStyle.NONE, size: 0, color: "auto"}, bottom: {style: BorderStyle.NONE, size: 0, color: "auto"}, left: {style: BorderStyle.NONE, size: 0, color: "auto"}, right: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideVertical: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideHorizontal: {style: BorderStyle.NONE, size: 0, color: "auto"} }
                    }),
                    new Paragraph({ text: "" }),
                    new Paragraph({ children: [ new TextRun({ text: "Kính gửi:  ", size: SZ_N, font: FONT }), new TextRun({ text: "Trưởng ban ban TTĐĐSX", bold: true, size: SZ_N, font: FONT }) ], alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: "" }),
                    new Paragraph({ children: [ new TextRun({ text: "ĐƠN ĐĂNG KÝ ĐI RA CÔNG TRÌNH BIỂN", bold: true, size: 28, font: FONT }) ], alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: "" }),
                    new Paragraph({ children: [new TextRun({ text: `Đơn vị đăng ký:      XNXL`, size: SZ_N, font: FONT })] }),
                    new Paragraph({ children: [new TextRun({ text: `Ngày khởi hành:  ${dateDisplayString}`, size: SZ_N, font: FONT })] }),
                    new Paragraph({ children: [new TextRun({ text: `Phương tiện yêu cầu:     Trực thăng x;     Tàu x`, size: SZ_N, font: FONT })] }),
                    new Paragraph({ children: [new TextRun({ text: `Đến CT biển : ${Object.keys(grouped).join(', ')}`, size: SZ_N, font: FONT })] }),
                    new Paragraph({ children: [new TextRun({ text: `Nhiệm vụ được giao: `, size: SZ_N, font: FONT })] })
                ];

                Object.keys(grouped).forEach(dest => {
                    const group = grouped[dest];
                    const meta = utils.getGroupFinalMeta(group);
                    const uniqueNVSX = meta.nvsxStr;
                    const uniqueTasks = (meta.taskLines || []);
                    children.push(new Paragraph({ children: [ new TextRun({ text: "-", size: SZ_N, font: FONT }), new TextRun({ text: ` ${dest}: `, bold: true, size: SZ_N, font: FONT }) ], indent: { left: 720 }, spacing: { before: 100 } }));
                    if(uniqueNVSX) children.push(new Paragraph({ children: [ new TextRun({ text: `+ NVSX: ${uniqueNVSX}`, size: SZ_N, font: FONT }) ], indent: { left: 1440 } }));
                    if(uniqueTasks.length > 0) {
                        children.push(new Paragraph({ children: [ new TextRun({ text: `+ Nội dung công việc:`, size: SZ_N, font: FONT }) ], indent: { left: 1440 } }));
                        uniqueTasks.forEach(line => { let l = (line || '').trim(); if (!l) return; if (/^\d+[\)\.]/.test(l)) l = l.replace(/^\d+[\)\.]\s*/, '- '); else if (!l.startsWith('-')) l = '- ' + l; children.push(new Paragraph({ children: [ new TextRun({ text: l, size: SZ_N, font: FONT }) ], indent: { left: 2160 } })); });
                    }
                });

                children.push(new Paragraph({ children: [new TextRun({ text: `Danh sách người đi: `, size: SZ_N, font: FONT })] }), new Paragraph({ text: "" }));

                const tableRows = [new TableRow({ children: ["Số TT", "Họ và tên", "Chức danh", "Đơn vị", "Danh số"].map(t => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, size: SZ_S, font: FONT })], alignment: AlignmentType.CENTER })], verticalAlign: "center", shading: { fill: "FFFFFF" } })), tableHeader: true })];

                Object.keys(grouped).forEach(dest => {
                    let dName = dest; if(!dest.toLowerCase().startsWith('giàn') && !dest.toLowerCase().startsWith('gian')) dName = `GIÀN ${dest}`;
                    
                    tableRows.push(new TableRow({ 
                        height: { value: 400, rule: HeightRule.AT_LEAST }, 
                        children: [
                            new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, shading: { fill: "D9D9D9" }, children: [] }), // Empty cell 1
                            new TableCell({ 
                                children: [
                                    new Paragraph({ 
                                        children: [new TextRun({ text: dName, bold: true, size: SZ_S, font: FONT })],
                                        alignment: AlignmentType.LEFT 
                                    })
                                ], 
                                columnSpan: 4, 
                                shading: { fill: "D9D9D9" },
                                verticalAlign: "center"
                            })
                        ] 
                    }));
                    
                    grouped[dest].forEach((p, i) => {
                        tableRows.push(new TableRow({
                            height: { value: 400, rule: HeightRule.AT_LEAST },
                            children: [
                                new TableCell({ children: [new Paragraph({ text: (i+1).toString(), alignment: AlignmentType.CENTER, size: SZ_S, font: FONT })], verticalAlign: "center" }),
                                new TableCell({ children: [new Paragraph({ text: p.fullName, size: SZ_S, font: FONT })], verticalAlign: "center" }),
                                new TableCell({ children: [new Paragraph({ text: p.title, size: SZ_S, font: FONT, alignment: AlignmentType.CENTER })], verticalAlign: "center" }),
                                new TableCell({ children: [new Paragraph({ text: p.orgName, size: SZ_S, font: FONT, alignment: AlignmentType.CENTER })], verticalAlign: "center" }),
                                new TableCell({ children: [new Paragraph({ text: p.staffNo, alignment: AlignmentType.CENTER, size: SZ_S, font: FONT })], verticalAlign: "center" }),
                            ]
                        }));
                    });
                });

                children.push(new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.SINGLE, size: 1 }, bottom: { style: BorderStyle.SINGLE, size: 1 }, left: { style: BorderStyle.SINGLE, size: 1 }, right: { style: BorderStyle.SINGLE, size: 1 }, insideVertical: { style: BorderStyle.SINGLE, size: 1 }, insideHorizontal: { style: BorderStyle.SINGLE, size: 1 } } }), new Paragraph({ text: "" }), new Paragraph({ text: "" }));

                children.push(new Table({
                    rows: [new TableRow({ children: [
                        new TableCell({ children: [
                            new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }),
                            new Paragraph({ children: [new TextRun({ text: "Ký tắt:", bold: true, italics: true, size: SZ_S, font: FONT })] }),
                            new Paragraph({ text: "" }), new Paragraph({ children: [new TextRun({ text: "- Lãnh đạo LDVN Vietsovpetro (nếu cần)", size: SZ_S, font: FONT })], indent: { left: 100 } }),
                            new Paragraph({ text: "" }), new Paragraph({ children: [new TextRun({ text: "- Điều độ – XNKT:", size: SZ_S, font: FONT })], indent: { left: 100 } }),
                            new Paragraph({ text: "" }), new Paragraph({ children: [new TextRun({ text: "- Phòng kỹ thuật :", size: SZ_S, font: FONT })], indent: { left: 100 } }),
                            new Paragraph({ text: "" }), new Paragraph({ children: [new TextRun({ text: "- Điều độ XNXL:                            Số đt: 8626", size: SZ_S, font: FONT })], indent: { left: 100 } }),
                        ], borders: { top: {style: BorderStyle.NONE, size: 0, color: "auto"}, bottom: {style: BorderStyle.NONE, size: 0, color: "auto"}, left: {style: BorderStyle.NONE, size: 0, color: "auto"}, right: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideVertical: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideHorizontal: {style: BorderStyle.NONE, size: 0, color: "auto"} }, width: { size: 50, type: WidthType.PERCENTAGE }, verticalAlign: "bottom" }),
                        new TableCell({ children: [
                            new Paragraph({ children: [new TextRun({ text: sig.approver || 'TRƯỞNG BAN TTĐĐSX', bold: true, size: SZ_S, font: FONT })], alignment: AlignmentType.CENTER }),
                            new Paragraph({ children: [new TextRun({ text: "(Ký, ghi rõ họ tên)", italics: true, size: 20, font: FONT })], alignment: AlignmentType.CENTER }),
                            new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }), 
                            new Paragraph({ children: [new TextRun({ text: sig.approverName || '', bold: true, size: SZ_S, font: FONT })], alignment: AlignmentType.CENTER })
                        ], borders: { top: {style: BorderStyle.NONE, size: 0, color: "auto"}, bottom: {style: BorderStyle.NONE, size: 0, color: "auto"}, left: {style: BorderStyle.NONE, size: 0, color: "auto"}, right: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideVertical: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideHorizontal: {style: BorderStyle.NONE, size: 0, color: "auto"} }, width: { size: 50, type: WidthType.PERCENTAGE }, verticalAlign: "top" })
                    ] })], width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: { top: {style: BorderStyle.NONE, size: 0, color: "auto"}, bottom: {style: BorderStyle.NONE, size: 0, color: "auto"}, left: {style: BorderStyle.NONE, size: 0, color: "auto"}, right: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideVertical: {style: BorderStyle.NONE, size: 0, color: "auto"}, insideHorizontal: {style: BorderStyle.NONE, size: 0, color: "auto"} }
                }));

                const blob = await Packer.toBlob(new Document({ sections: [{ properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children }] }));
                saveAs(blob, `KeHoachDiBien_${dateDisplayString.split('+')[0].trim().replace(/\./g,'_')}_plus.docx`);
            },
                        generatePdf: async (peopleData, dateDisplayString) => {
                const canAutoPdf = (typeof html2pdf !== 'undefined');
                // If html2pdf isn't available (offline/CDN blocked), we will fallback to print-to-PDF.

                let sig = { approver: 'TRƯỞNG BAN TTĐĐSX', approverName: '' };
                try {
                    const sigDoc = await getPublicColl('settings').doc('signature').get();
                    if (sigDoc.exists) sig = sigDoc.data();
                } catch(e) { /* ignore */ }

                const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
                const today = new Date();
                const dd = String(today.getDate()).padStart(2,'0');
                const mm = String(today.getMonth()+1).padStart(2,'0');
                const yyyy = today.getFullYear();

                const people = [...peopleData];
                people.sort((a, b) => (a.destination||'').localeCompare(b.destination||'') || (a.nvsxNo||'').localeCompare(b.nvsxNo||'') || (a.orgName||'').localeCompare(b.orgName||''));
                const grouped = people.reduce((acc, p) => { const k = p.destination || 'KHÁC'; (acc[k] = acc[k] || []).push(p); return acc; }, {});

                const destKeys = Object.keys(grouped);
                const destList = destKeys.join(', ');

                // Build nhiệm vụ section
                let tasksHtml = '';
                destKeys.forEach(dest => {
                    const group = grouped[dest];
                    const meta = utils.getGroupFinalMeta(group);
                    const nvsx = meta.nvsxStr;
                    const lines = meta.taskLines || [];
                    tasksHtml += `
                        <div style="margin-left: 10mm; margin-top: 2mm;">
                            <div><span style="font-weight:bold;">- ${esc(dest)}:</span></div>
                            ${nvsx ? `<div style="margin-left: 6mm;">+ NVSX: ${esc(nvsx)}</div>` : ``}
                            ${lines.length ? `<div style="margin-left: 6mm;">+ Nội dung công việc:</div>` : ``}
                            ${lines.map(l => {
                                let x = (l||'').trim();
                                if (!x) return '';
                                if (/^\d+[\)\.]/.test(x)) x = x.replace(/^\d+[\)\.]\s*/, '- ');
                                else if (!x.startsWith('-')) x = '- ' + x;
                                return `<div style="margin-left: 12mm;">${esc(x)}</div>`;
                            }).join('')}
                        </div>
                    `;
                });

                // Build people table
                let tableRows = `
                    <tr>
                        <th style="border:1px solid #000; padding:3mm; width:10mm; text-align:center;">Số TT</th>
                        <th style="border:1px solid #000; padding:3mm; text-align:center;">Họ và tên</th>
                        <th style="border:1px solid #000; padding:3mm; text-align:center;">Chức danh</th>
                        <th style="border:1px solid #000; padding:3mm; text-align:center;">Đơn vị</th>
                        <th style="border:1px solid #000; padding:3mm; width:22mm; text-align:center;">Danh số</th>
                    </tr>
                `;

                destKeys.forEach(dest => {
                    let dName = dest;
                    const dLower = String(dest || '').toLowerCase();
                    if (!dLower.startsWith('giàn') && !dLower.startsWith('gian')) dName = `GIÀN ${dest}`;
                    tableRows += `
                        <tr>
                            <td style="border:1px solid #000; padding:3mm; background:#D9D9D9;"></td>
                            <td style="border:1px solid #000; padding:3mm; background:#D9D9D9; font-weight:bold;" colspan="4">${esc(dName)}</td>
                        </tr>
                    `;
                    grouped[dest].forEach((p, i) => {
                        tableRows += `
                            <tr>
                                <td style="border:1px solid #000; padding:2.5mm; text-align:center;">${i+1}</td>
                                <td style="border:1px solid #000; padding:2.5mm;">${esc(p.fullName||'')}</td>
                                <td style="border:1px solid #000; padding:2.5mm; text-align:center;">${esc(p.title||'')}</td>
                                <td style="border:1px solid #000; padding:2.5mm; text-align:center;">${esc(p.orgName||'')}</td>
                                <td style="border:1px solid #000; padding:2.5mm; text-align:center;">${esc(p.staffNo||'')}</td>
                            </tr>
                        `;
                    });
                });

                const root = document.getElementById('pdfExportRoot') || (() => {
                    const d = document.createElement('div'); d.id = 'pdfExportRoot'; d.className = 'hidden'; document.body.appendChild(d); return d;
                })();

                root.classList.remove('hidden');
                root.innerHTML = `
                    <div id="pdfDoc" style="font-family: 'Times New Roman', Times, serif; font-size: 12pt; color: #000; line-height: 1.25;">
                        <style>
                            @page { size: A4; margin: 12.7mm; }
                        </style>

                        <table style="width:100%; border-collapse:collapse; margin-bottom:6mm;">
                            <tr>
                                <td style="width:40%; vertical-align:top; text-align:center;">
                                    <div style="font-weight:bold;">LIÊN DOANH VIỆT – NGA</div>
                                    <div style="font-weight:bold;">VIETSOVPETRO</div>
                                    <div style="font-weight:bold;">XNXL</div>
                                    <div style="height:6mm;"></div>
                                    <div>Số: ...../.....-CV-XL</div>
                                </td>
                                <td style="width:60%; vertical-align:top; text-align:center;">
                                    <div style="font-weight:bold;">CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
                                    <div style="font-weight:bold; text-decoration: underline;">Độc lập – Tự do – Hạnh phúc</div>
                                    <div style="height:6mm;"></div>
                                    <div style="font-style:italic;">TPHCM, ngày ${dd} tháng ${mm} năm ${yyyy}</div>
                                </td>
                            </tr>
                        </table>

                        <div style="text-align:center; margin-bottom:4mm;">
                            <span>Kính gửi: </span><span style="font-weight:bold;">Trưởng ban ban TTĐĐSX</span>
                        </div>

                        <div style="text-align:center; font-weight:bold; font-size:14pt; margin: 4mm 0 6mm 0;">ĐƠN ĐĂNG KÝ ĐI RA CÔNG TRÌNH BIỂN</div>

                        <div>Đơn vị đăng ký:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;XNXL</div>
                        <div>Ngày khởi hành:&nbsp;&nbsp;${esc(dateDisplayString)}</div>
                        <div>Phương tiện yêu cầu:&nbsp;&nbsp;&nbsp;&nbsp;Trực thăng x;&nbsp;&nbsp;&nbsp;&nbsp;Tàu x</div>
                        <div>Đến CT biển : ${esc(destList)}</div>
                        <div style="margin-top:2mm;">Nhiệm vụ được giao:</div>

                        ${tasksHtml}

                        <div style="margin-top:4mm;">Danh sách người đi:</div>
                        <div style="height:2mm;"></div>

                        <table style="width:100%; border-collapse:collapse;">
                            ${tableRows}
                        </table>

                        <div style="height:8mm;"></div>

                        <table style="width:100%; border-collapse:collapse;">
                            <tr>
                                <td style="width:50%; vertical-align:bottom;">
                                    <div style="height:26mm;"></div>
                                    <div style="font-style:italic; font-weight:bold;">Ký tắt:</div>
                                    <div style="height:2mm;"></div>
                                    <div style="margin-left:3mm;">- Lãnh đạo LDVN Vietsovpetro (nếu cần)</div>
                                    <div style="height:2mm;"></div>
                                    <div style="margin-left:3mm;">- Điều độ – XNKT:</div>
                                    <div style="height:2mm;"></div>
                                    <div style="margin-left:3mm;">- Phòng kỹ thuật :</div>
                                    <div style="height:2mm;"></div>
                                    <div style="margin-left:3mm;">- Điều độ XNXL:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Số đt: 8626</div>
                                </td>
                                <td style="width:50%; vertical-align:top; text-align:center;">
                                    <div style="font-weight:bold;">${esc(sig.approver || 'TRƯỞNG BAN TTĐĐSX')}</div>
                                    <div style="font-style:italic; font-size:10pt;">(Ký, ghi rõ họ tên)</div>
                                    <div style="height:28mm;"></div>
                                    <div style="font-weight:bold;">${esc(sig.approverName || '')}</div>
                                </td>
                            </tr>
                        </table>
                    </div>
                `;

                const filename = `KeHoachDiBien_${dateDisplayString.split('+')[0].trim().replace(/\./g,'_')}_plus.pdf`;

                try {
                    if (canAutoPdf) {
                        const opt = {
                            margin: 12.7,
                            filename,
                            image: { type: 'jpeg', quality: 0.98 },
                            html2canvas: { scale: 2, useCORS: true },
                            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                            pagebreak: { mode: ['css', 'legacy'] }
                        };
                        await html2pdf().set(opt).from(root.querySelector('#pdfDoc')).save();
                    } else {
                        // Fallback: open a print window so user can Save as PDF.
                        const w = window.open('', '_blank');
                        if (!w) throw new Error('Không mở được cửa sổ in PDF (popup bị chặn).');
                        const html = `<!doctype html><html><head><meta charset='utf-8'><title>${filename}</title></head><body>${root.querySelector('#pdfDoc').outerHTML}</body></html>`;
                        w.document.open();
                        w.document.write(html);
                        w.document.close();
                        setTimeout(() => { try { w.focus(); w.print(); } catch(e) {} }, 300);
                        utils.showToast('Không có thư viện PDF tự động. Đã mở cửa sổ in — chọn "Save as PDF" để lưu.', 'info');
                    }
                } finally {
                    root.innerHTML = '';
                    root.classList.add('hidden');
                }
            }
        };

        app.init();

// --- Expose handlers for inline onclick (safety) ---
try { window.app = app; } catch(e) { /* ignore */ }
try { window.auditModal = auditModal; } catch(e) { /* ignore */ }
try { window.editModal = editModal; } catch(e) { /* ignore */ }
try { window.importModal = importModal; } catch(e) { /* ignore */ }
try { window.exportManager = exportManager; } catch(e) { /* ignore */ }
try { window.reportManager = reportManager; } catch(e) { /* ignore */ }
try { window.userModal = userModal; } catch(e) { /* ignore */ }
try { window.warningCenter = warningCenter; } catch(e) { /* ignore */ }
try { window.masterDataManager = masterDataManager; } catch(e) { /* ignore */ }
try { window.userManager = userManager; } catch(e) { /* ignore */ }
