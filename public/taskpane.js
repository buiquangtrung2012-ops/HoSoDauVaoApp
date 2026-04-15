import { WordService } from './word_service.js';
import { StorageService } from './storage_service.js';
import { MockData } from './mock_data.js';

// --- GLOBAL STATE ---
let state = {
    currentTab: 'duAn',
    editingIndex: -1,
    templateFile: null, // Lưu trữ file mẫu người dùng chọn
    templateFileName: "",
    duAn: {
        soHD: "",
        tenDuAn: "",
        goiThau: "",
        dvtc: "",
        daiDienCDT: "",
        tvgs: "",
        ngayKhoiCong: "",
        ngayHoanThanh: ""
    },
    nhanSu: [],
    mayMoc: [],
    vatLieu: [],
    thiNghiem: [],
    outputMode: 'multiple'
};

const categories = {
    duAn: { title: "Dự án", fields: ["tenDuAn", "goiThau", "dvtc", "daiDienCDT", "tvgs", "soHD", "ngayKhoiCong", "ngayHoanThanh"], labels: ["Tên dự án", "Tên gói thầu", "Đơn vị thi công", "Đại diện CDT", "Tư vấn giám sát", "Số hợp đồng", "Ngày khởi công", "Ngày hoàn thành"] },
    nhanSu: { title: "Nhân sự", fields: ["stt", "name", "role", "major", "phone"], labels: ["STT", "Họ và tên", "Chức danh", "Chuyên ngành", "Số điện thoại"] },
    mayMoc: { title: "Máy móc", fields: ["stt", "name", "unit", "qty", "owner", "status"], labels: ["STT", "Tên thiết bị", "Đơn vị tính", "Số lượng", "Chủ sở hữu", "Hình thức"] },
    vatLieu: { title: "Vật liệu", fields: ["stt", "name", "standard", "origin", "note"], labels: ["STT", "Tên vật tư", "Thông số/Tiêu chuẩn", "Nguồn gốc", "Đơn vị cung cấp"] },
    thiNghiem: { title: "Phòng TN", fields: ["stt", "dvtn", "address", "ptn", "func"], labels: ["STT", "Đơn vị TN", "Địa chỉ", "Tên phòng TN", "Chức năng"] }
};

// --- INITIALIZATION ---
window.onload = async () => {
    await initializeApp();
};

async function initializeApp() {
    try {
        console.log("Desktop App Initializing...");
        await loadState();
        registerEvents();
        switchTab('duAn');
        if (typeof lucide !== 'undefined') lucide.createIcons();
    } catch (e) {
        console.error("Initialization error:", e);
    }
}

async function loadState() {
    for (const key of ['duAn', 'nhanSu', 'mayMoc', 'vatLieu', 'thiNghiem']) {
        const saved = await StorageService.getProjectData(key);
        if (saved) state[key] = saved;
    }
    const savedTemplateName = localStorage.getItem('hoso_templateName');
    if (savedTemplateName) {
        state.templateFileName = savedTemplateName;
        updateTemplateLabel();
    }
}

async function saveState() {
    for (const key of ['duAn', 'nhanSu', 'mayMoc', 'vatLieu', 'thiNghiem']) {
        await StorageService.setProjectData(key, state[key]);
    }
}

function updateTemplateLabel() {
    const label = document.getElementById('templateLabel');
    if (label) {
        label.innerText = state.templateFileName ? `Mẫu: ${state.templateFileName}` : "Chưa chọn file mẫu .docx";
        label.classList.toggle('text-indigo-600', !!state.templateFileName);
        label.classList.toggle('font-bold', !!state.templateFileName);
    }
}

function registerEvents() {
    document.querySelectorAll('[data-tab]').forEach(btn => {
        btn.onclick = () => switchTab(btn.dataset.tab);
    });

    // Xử lý chọn file mẫu
    const fileInput = document.getElementById('templateInput');
    if (fileInput) {
        fileInput.onchange = (e) => {
            const file = e.target.files[0];
            if (file) {
                state.templateFile = file;
                state.templateFileName = file.name;
                localStorage.setItem('hoso_templateName', file.name);
                updateTemplateLabel();
                showToast("Đã tải file mẫu thành công!", "success");
            }
        };
    }

    document.getElementById('btnCapNhat').onclick = async () => {
        await saveState();
        if (!state.templateFile) {
            showToast("Vui lòng chọn file mẫu .docx trước!", "warning");
            return;
        }

        try {
            updateLog("Đang xử lý điền dữ liệu vào file mẫu...");
            const resultBlob = await WordService.fillTemplate(state.templateFile, state);
            if (resultBlob) {
                const downloadName = `HoSo_${state.duAn.tenDuAn || 'Moi'}.docx`.replace(/\s+/g, '_');
                saveAs(resultBlob, downloadName);
                showToast("Đã xuất file Word thành công!", "success");
            }
        } catch (err) {
            console.error(err);
            showToast("Lỗi khi điền dữ liệu: " + err.message, "error");
        }
    };

    document.getElementById('btnExport').onclick = () => {
        const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(state));
        const downloadAnchorNode = document.createElement('a');
        downloadAnchorNode.setAttribute("href", dataStr);
        downloadAnchorNode.setAttribute("download", "hoso_data.json");
        document.body.appendChild(downloadAnchorNode);
        downloadAnchorNode.click();
        downloadAnchorNode.remove();
    };

    // Nạp dữ liệu mẫu
    document.getElementById('btnImportDoc').onclick = async () => {
        if (confirm("Bạn có muốn nạp dữ liệu mẫu để dùng thử?")) {
            state.duAn = { ...MockData.duAn };
            state.nhanSu = [...MockData.nhanSu];
            state.mayMoc = [...MockData.mayMoc];
            state.vatLieu = [...MockData.vatLieu];
            state.thiNghiem = [...MockData.thiNghiem];
            await saveState();
            renderContent();
            showToast("Đã nạp dữ liệu mẫu!", "success");
        }
    };
}

function switchTab(tabId) {
    state.currentTab = tabId;
    document.querySelectorAll('[data-tab]').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });
    renderContent();
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function renderContent() {
    const container = document.getElementById('tabContent');
    if (!container) return;
    container.innerHTML = "";
    
    if (state.currentTab === 'duAn') {
        renderProjectForm(container);
    } else {
        renderList(container, state.currentTab);
    }
}

function renderProjectForm(container) {
    const form = document.createElement("div");
    form.className = "project-grid-container bg-white p-6 rounded-2xl border border-slate-100 shadow-sm";
    
    categories.duAn.fields.forEach((field, i) => {
        const div = document.createElement("div");
        div.innerHTML = `
            <label class="text-[10px] font-black text-slate-400 uppercase mb-1 ml-1">${categories.duAn.labels[i]}</label>
            <textarea data-field="${field}" class="input-field project-input resize-none py-2" rows="1">${state.duAn[field] || ''}</textarea>
        `;
        form.appendChild(div);
    });
    
    container.appendChild(form);
    form.querySelectorAll('.project-input').forEach(input => {
        input.onchange = () => {
            state.duAn[input.dataset.field] = input.value;
            saveState();
        };
    });
}

function renderList(container, type) {
    const items = state[type] || [];
    const config = categories[type];
    
    const card = document.createElement('div');
    card.className = 'bg-white p-4 rounded-2xl border border-slate-100 shadow-sm';
    
    let headersHtml = config.labels.map(l => `<th class="border p-2 text-left bg-slate-50 text-[10px]">${l}</th>`).join("");
    
    card.innerHTML = `
        <div class="flex justify-between items-center mb-4">
            <h3 class="font-bold text-slate-700">${config.title}</h3>
            <button id="btnAddRow" class="h-8 px-3 bg-indigo-600 text-white rounded-lg text-xs font-bold hover:bg-indigo-700">
                Thêm dòng mới
            </button>
        </div>
        <div class="overflow-x-auto">
            <table class="w-full border-collapse text-[11px]">
                <thead><tr class="text-slate-500 uppercase font-bold">${headersHtml}</tr></thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    `;
    
    const tbody = card.querySelector('#tableBody');
    items.forEach((item, idx) => {
        const tr = document.createElement('tr');
        config.fields.forEach((f, i) => {
            const td = document.createElement('td');
            td.className = 'border p-1';
            const input = document.createElement('input');
            input.type = 'text';
            input.value = item[i] || '';
            input.className = 'w-full p-1 border-none focus:ring-0 bg-transparent';
            input.onchange = (e) => {
                state[type][idx][i] = e.target.value;
                saveState();
            };
            td.appendChild(input);
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    
    container.appendChild(card);
    card.querySelector('#btnAddRow').onclick = () => {
        const newRow = new Array(config.fields.length).fill("");
        newRow[0] = (state[type].length + 1).toString();
        state[type].push(newRow);
        renderContent();
    };
}

function updateLog(msg) {
    console.log(msg);
}

function showToast(msg, type = "success") {
    // Tạm thời dùng alert hoặc một div toast đơn giản
    const toast = document.createElement('div');
    toast.className = `fixed bottom-4 right-4 px-6 py-3 rounded-xl text-white font-bold z-[100] shadow-2xl transition-all ${type === 'success' ? 'bg-emerald-500' : 'bg-amber-500'}`;
    toast.innerText = msg;
    document.body.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 500);
    }, 3000);
}
