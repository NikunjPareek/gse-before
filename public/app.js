const tableDefaults = [
    { component: 'PV Modules',                              spec: 'BIFICAL',                              company: 'Adani',            qty: '10 No.'      },
    { component: 'Grid-Tie Inverter',                        spec: 'Capacity of {{kw}} kW',                company: 'K-Solar/UTL',      qty: '1 Nos.'      },
    { component: 'Structure',                                spec: 'GP Pipe (Apollo) Galvanized',          company: 'Apollo 2mm',       qty: 'As Required' },
    { component: 'Meter',                                    spec: 'Net Meter & Solar',                    company: 'Genius / L & T',   qty: '1 set'       },
    { component: 'DC Wire',                                  spec: '4 mm',                                 company: 'Polycab',          qty: 'As Required' },
    { component: 'AC Wire',                                  spec: 'As Required',                          company: 'Aluminium Armoured',qty: 'As Required' },
    { component: 'Earthing Wire',                            spec: '1" Strip',                             company: 'GI',               qty: 'As Required' },
    { component: 'Earthing Material',                        spec: '2 Meter GI Earthing Rod + Chemical',   company: 'As Per Standard',  qty: '1 Set'       },
    { component: 'Lighting Arrestor',                        spec: 'As Required',                          company: 'As Per Standard',  qty: '1 No.'       },
    { component: 'Fitting Accessories',                      spec: 'As Required',                          company: 'As Per Standard',  qty: 'As Required' },
    { component: 'ACDB / DCDB / MCB Box / Busbar / Panel Box', spec: 'As Required',                       company: 'As Per Standard',  qty: '1 Set'       }
];

let globalCount = 0;

function getFormattedDate() {
  const inputDate = document.getElementById('date').value;

  let dateObj;
  if (inputDate) {
    dateObj = new Date(inputDate); 
  } else {
    dateObj = new Date(); 
  }
  const dd = String(dateObj.getDate()).padStart(2, '0');
  const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  const yy = String(dateObj.getFullYear()).slice(-2);

  return `${dd}${mm}${yy}`;
}

function getClientShortName() {
    const name = document.getElementById('client_name').value.trim();
    if (!name) return 'NA';

    const words = name.split(' ');

    if (words.length === 1) {
        return words[0].slice(0, 3).toUpperCase();
    }

    const first = words[0].slice(0, 3);
    const last = words[words.length - 1].slice(0, 3);

    return (first + last).toUpperCase();
}

function convertNumberToWords(amount) {
    const words = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
      'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    
    if (amount === 0) return 'Zero Rupees Only';
    let a = Math.round(amount).toString();
    if (a.length > 9) return 'Amount too large';
  
    let n = ('000000000' + a).slice(-9).match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
    if (!n) return '';
    let str = '';
    str += (n[1] != 0) ? (words[Number(n[1])] || tens[n[1][0]] + ' ' + words[n[1][1]]) + ' Crore ' : '';
    str += (n[2] != 0) ? (words[Number(n[2])] || tens[n[2][0]] + ' ' + words[n[2][1]]) + ' Lakh ' : '';
    str += (n[3] != 0) ? (words[Number(n[3])] || tens[n[3][0]] + ' ' + words[n[3][1]]) + ' Thousand ' : '';
    str += (n[4] != 0) ? (words[Number(n[4])] || tens[n[4][0]] + ' ' + words[n[4][1]]) + ' Hundred ' : '';
    str += (n[5] != 0) ? ((str != '') ? 'and ' : '') + (words[Number(n[5])] || tens[n[5][0]] + ' ' + words[n[5][1]]) + ' ' : '';
    return str.trim() + ' Rupees Only';
}

function formatCurrency(amount) {
    return new Intl.NumberFormat('en-IN', { maximumFractionDigits: 2 }).format(amount);
}

function renderTable() {
    const tbody = document.getElementById('materials-body');
    tbody.innerHTML = '';
    for (let i = 0; i < 11; i++) {
        const d = tableDefaults[i];
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${i + 1}</td>
            <td>${d.component}</td>
            <td><input type="text" id="spec_${i + 1}" class="row-input"></td>
            <td><input type="text" id="company_${i + 1}" class="row-input"></td>
            <td><input type="text" id="qty_${i + 1}" class="row-input"></td>
        `;
        tbody.appendChild(tr);
    }
}

function loadDefaults() {
    const kw = document.getElementById('kw').value || '';
    for (let i = 0; i < 11; i++) {
        const d = tableDefaults[i];
        let specStr = d.spec;
        // For inverter row, substitute {{kw}} with the current kW value
        if (i === 1) specStr = specStr.replace('{{kw}}', kw);
        document.getElementById(`spec_${i + 1}`).value = specStr;
        document.getElementById(`company_${i + 1}`).value = d.company;
        document.getElementById(`qty_${i + 1}`).value = d.qty;
    }
}

function validateInputs() {
    const clientName = document.getElementById('client_name').value.trim();
    const kw = document.getElementById('kw').value.trim();
    const baseCost = document.getElementById('base_cost').value.trim();
    
    if(!clientName || !kw || !baseCost) {
        showToast('Please fill out Client Name, Capacity (kW), and Base Cost.', 'error');
        return false;
    }
    return true;
}

function calculate() {
    const type = document.querySelector('input[name="type"]:checked').value;
    const base_cost = parseFloat(document.getElementById('base_cost').value) || 0;
    const gst_pct = parseFloat(document.getElementById('gst_pct').value) || 0;
    const central_subsidy = parseFloat(document.getElementById('central_subsidy').value) || 0;
    const state_subsidy = parseFloat(document.getElementById('state_subsidy').value) || 0;
    const discount = parseFloat(document.getElementById('discount_amount').value) || 0;
    
    document.getElementById('discount-container').style.display = type === 'client' ? 'block' : 'none';

    const gst_amount = (base_cost * gst_pct) / 100;
    const total_subsidy = central_subsidy + state_subsidy;
    
    let final_amount = base_cost + gst_amount;
    if (type === 'client') {
        final_amount -= discount;
    }

    document.getElementById('gst_display').value = formatCurrency(gst_amount);
    document.getElementById('total_subsidy').value = formatCurrency(total_subsidy);
    document.getElementById('final_amount').value = formatCurrency(final_amount);
    document.getElementById('final_amount_words').value = convertNumberToWords(final_amount);
    
    updateRefNo();
}

function updateRefNo(){ 
    const type = document.querySelector('input[name="type"]:checked').value;
    const kw = document.getElementById('kw').value || '0';
    const padded = globalCount.toString().padStart(4, '0');
    const tStr = type === 'bank' ? 'B' : 'C';
    const dateStr = getFormattedDate(); 
    const clientShort = getClientShortName();
    document.getElementById('ref_no').value = `GSE/${tStr}/${kw}kW/${dateStr}/${clientShort}/${padded}`;
}

function showToast(msg, type) {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = msg;
    container.appendChild(toast);
    setTimeout(() => {
        toast.remove();
    }, 4000);
}

function setStatus(state) {
    const dot = document.getElementById('status-dot');
    const text = document.getElementById('status-text');
    dot.className = `dot ${state}`;
    if (state === 'idle') text.textContent = 'Ready';
    if (state === 'loading') text.textContent = 'Generating...';
    if (state === 'success') text.textContent = 'Done';
    if (state === 'error') text.textContent = 'Failed';
}

function buildPayload() {
    // Only send exact placeholders expected by templates
    const client_number = document.getElementById('client_number').value;
    const base_cost = parseFloat(document.getElementById('base_cost').value) || 0;
    const gst_pct = parseFloat(document.getElementById('gst_pct').value) || 0;
    const gst = (base_cost * gst_pct) / 100;
    const central_subsidy = parseFloat(document.getElementById('central_subsidy').value) || 0;
    const state_subsidy = parseFloat(document.getElementById('state_subsidy').value) || 0;
    const total_subsidy = central_subsidy + state_subsidy;
    const type = document.querySelector('input[name="type"]:checked').value;
    let discount = 0;
    if (type === 'client') {
        discount = parseFloat(document.getElementById('discount_amount').value) || 0;
    }
    
    let final_amount = base_cost + gst - total_subsidy - discount;

    const payload = {
        base_cost: formatCurrency(base_cost),
        client_address: document.getElementById('client_address').value || '',
        client_name: document.getElementById('client_name').value || '',
        date: document.getElementById('date').value ? new Date(document.getElementById('date').value).toLocaleDateString('en-GB') : '',
        final_amount_words: convertNumberToWords(final_amount),
        final_amount: formatCurrency(final_amount),
        gst: formatCurrency(gst),
        kw: document.getElementById('kw').value || '0',
        ref_no: document.getElementById('ref_no').value || ''
    };
    
    if (type === 'client') {
        payload['Discount_amount'] = formatCurrency(discount);
    }
    
    // spec_1..11 and company_1..11 all exist in templates
    // qty_1, qty_2, qty_4 are the only qty placeholders in the templates
    for (let i = 1; i <= 11; i++) {
        payload[`spec_${i}`] = document.getElementById(`spec_${i}`).value || '';
        payload[`company_${i}`] = document.getElementById(`company_${i}`).value || '';
    }
    payload['qty_1'] = document.getElementById('qty_1').value || '';
    payload['qty_2'] = document.getElementById('qty_2').value || '';
    payload['qty_4'] = document.getElementById('qty_4').value || '';
    return payload;
}

function saveState() {
    const state = {
        type: document.querySelector('input[name="type"]:checked').value,
        date: document.getElementById('date').value,
        kw: document.getElementById('kw').value,
        client_name: document.getElementById('client_name').value,
        client_address: document.getElementById('client_address').value,
        base_cost: document.getElementById('base_cost').value,
        gst_pct: document.getElementById('gst_pct').value,
        central_subsidy: document.getElementById('central_subsidy').value,
        state_subsidy: document.getElementById('state_subsidy').value,
        discount_amount: document.getElementById('discount_amount').value,
        table: {}
    };
    for(let i=1; i<=11; i++) {
        state.table[i] = {
            spec: document.getElementById(`spec_${i}`).value,
            company: document.getElementById(`company_${i}`).value,
            qty: document.getElementById(`qty_${i}`).value
        };
    }
    localStorage.setItem('gse_quote_state', JSON.stringify(state));
}

function restoreState() {
    const saved = localStorage.getItem('gse_quote_state');
    if (saved) {
        try {
            const state = JSON.parse(saved);
            document.querySelector(`input[name="type"][value="${state.type}"]`).checked = true;
            document.getElementById('date').value = state.date;
            document.getElementById('kw').value = state.kw;
            document.getElementById('client_name').value = state.client_name;
            document.getElementById('client_address').value = state.client_address;
            document.getElementById('base_cost').value = state.base_cost;
            document.getElementById('gst_pct').value = state.gst_pct;
            document.getElementById('central_subsidy').value = state.central_subsidy;
            document.getElementById('state_subsidy').value = state.state_subsidy;
            document.getElementById('discount_amount').value = state.discount_amount;
            
            for(let i=1; i<=11; i++) {
                if(state.table[i]) {
                    document.getElementById(`spec_${i}`).value = state.table[i].spec;
                    document.getElementById(`company_${i}`).value = state.table[i].company;
                    document.getElementById(`qty_${i}`).value = state.table[i].qty;
                }
            }
        } catch(e) {}
    } else {
        loadDefaults();
    }
}

async function fetchCount() {
    try {
        const res = await fetch('/api/next-count');
        const data = await res.json();
        globalCount = data.count;
        updateRefNo();
    } catch(e) { console.error(e); }
}

async function incrementCount() {
    try {
        const res = await fetch('/api/increment-count', { method: 'POST' });
        const data = await res.json();
        globalCount = data.count;
        updateRefNo();
    } catch(e) { console.error(e); }
}

async function downloadDoc(ext) {
    if(!validateInputs()) return;
    setStatus('loading');
    
    const type = document.querySelector('input[name="type"]:checked').value;
    const payload = buildPayload();
    
    try {
        const res = await fetch(`/api/download/${ext}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ type, data: payload })
        });
        
        if(!res.ok) {
            const err = await res.json();
            throw new Error(err.error || 'Failed to download');
        }
        
        let filename = res.headers.get('content-disposition');
        if(filename && filename.indexOf('filename=') !== -1) {
            filename = filename.split('filename=')[1].replace(/"/g, '');
        } else {
            filename = `Quotation.${ext}`;
        }
        
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.hidden = true;
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
        
        setStatus('success');
        showToast(`Successfully downloaded ${ext.toUpperCase()}`, 'success');
        incrementCount();
        
    } catch (error) {
        console.error(error);
        setStatus('error');
        showToast(error.message, 'error');
    }
}

function init() {
    renderTable();
    restoreState();
    fetchCount();
    calculate();
    
    // Attach listeners
    document.querySelectorAll('input, textarea').forEach(el => {
        el.addEventListener('input', () => {
            if (el.id === 'kw') {
                const spec2 = document.getElementById('spec_2');
                if (spec2.value.startsWith('Capacity of')) {
                    spec2.value = `Capacity of ${el.value} kW`;
                }
            }
            calculate();
            updateRefNo();   
            saveState();
        });
    });

    document.getElementById('date').addEventListener('change', () => {
        updateRefNo();
        saveState();
    });

    document.querySelectorAll('input[name="type"]').forEach(el => {
        el.addEventListener('change', () => {
            calculate();
            updateRefNo();
            saveState();
        });
    });

    document.getElementById('btn-doc').addEventListener('click', () => downloadDoc('docx'));
    document.getElementById('btn-pdf').addEventListener('click', () => downloadDoc('pdf'));

    document.getElementById('clear-btn').addEventListener('click', () => {
        localStorage.removeItem('gse_quote_state');
        document.querySelectorAll('input:not([type="radio"]), textarea').forEach(el => el.value = '');
        document.getElementById('gst_pct').value = '8.9';
        document.getElementById('central_subsidy').value = '78000';
        document.getElementById('state_subsidy').value = '17000';
        document.querySelector('input[name="type"][value="bank"]').checked = true;
        loadDefaults();
        fetchCount();
        calculate();
    });
}

document.addEventListener('DOMContentLoaded', init);