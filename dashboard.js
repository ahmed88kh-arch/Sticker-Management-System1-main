// ==================== EXCEL EXPORT (SheetJS) ====================
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const os = require('os');

// Safe Electron dialog loader (won't crash if remote is unavailable)
let electronDialog = null;
try { electronDialog = require('electron').remote.dialog; } catch (e1) {
    try { electronDialog = require('@electron/remote').dialog; } catch (e2) {
        electronDialog = null; // Will fallback to Desktop save
    }
}

// ==================== DOM ELEMENTS ====================
// Stickers
const stickerModal = document.getElementById('stickerModal');
const openModalBtn = document.getElementById('openModalBtn');
const closeModalBtn = document.getElementById('closeModalBtn');
const stickerForm = document.getElementById('stickerForm');
const tableBody = document.getElementById('stickerTableBody');
const pasteArea = document.getElementById('apPasteArea');
const pastePreview = document.getElementById('apPastePreview');

// Workers
const workerModal = document.getElementById('workerModal');
const openWorkerModalBtn = document.getElementById('openWorkerModalBtn');
const closeWorkerModalBtn = document.getElementById('closeWorkerModalBtn');
const cancelWorkerBtn = document.getElementById('cancelWorkerBtn');
const workerForm = document.getElementById('workerForm');
const workerTableBody = document.getElementById('workerTableBody');

// Sales
const addToCartForm = document.getElementById('addToCartForm');
let currentInvoiceCart = [];
const salesProductSelect = document.getElementById('salesProductSelect');
const salesTableBody = document.getElementById('salesTableBody');
const workerAssignmentList = document.getElementById('workerAssignmentList');
const addWorkerToSaleBtn = document.getElementById('addWorkerToSaleBtn');

// Inventory (Handled inside custom section)

// Reports
const workerReportTableBody = document.getElementById('workerReportTableBody');

// Navigation
const navItems = document.querySelectorAll('.nav-item');
const pageSections = document.querySelectorAll('.page-section');

// ==================== NAVIGATION LOGIC ====================
window.forceGoHome = function() {
    document.querySelectorAll('.page-section').forEach(section => {
        section.style.display = section.id === 'sectionHomeMenu' ? 'block' : 'none';
    });
    
    const topBarRight = document.querySelector('.top-bar-right');
    const sidebarHeader = document.querySelector('.sidebar-header');
    if (topBarRight) topBarRight.style.display = 'flex';
    if (sidebarHeader) sidebarHeader.style.display = 'flex';
    
    document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
    const homeBtn = document.querySelector('.home-btn');
    if (homeBtn) homeBtn.classList.add('active');
    
    try {
        if (typeof updateStats === 'function') updateStats();
    } catch(e) { console.error('forceGoHome error', e); }
};
// Use event delegation on the document for reliability
document.addEventListener('click', (e) => {
    // Find the clicked nav-item (even if click was on icon/text inside)
    const navItem = e.target.closest('.nav-item');
    if (!navItem) return;
    
    e.preventDefault();
    console.log('Navigation clicked: ', navItem.dataset.target);
    const target = navItem.getAttribute('data-target');
    if (!target) return;

    // Update active class
    document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
    navItem.classList.add('active');

    // Show target section, hide others
    document.querySelectorAll('.page-section').forEach(section => {
        section.style.display = section.id === target ? 'block' : 'none';
    });

    // Hide top bar elements when not in the home dashboard
    const topBarRight = document.querySelector('.top-bar-right');
    const sidebarHeader = document.querySelector('.sidebar-header');
    
    if (topBarRight && sidebarHeader) {
        if (target === 'sectionHomeMenu') {
            topBarRight.style.display = 'flex';
            sidebarHeader.style.display = 'flex';
        } else {
            topBarRight.style.display = 'none';
            sidebarHeader.style.display = 'none';
        }
    }

    // Run tab-specific rendering/refresh logic to keep UI in sync
    if (target === 'sectionHomeMenu') {
        try {
            updateStats();
        } catch (e) {
            console.error('sectionHomeMenu init error', e);
        }
    }
    if (target === 'sectionDashboard') {
        try {
            renderTable();
            updateStats();
        } catch (e) {
            console.error('sectionDashboard init error', e);
        }
    }
    if (target === 'sectionWorkers') {
        try {
            renderWorkerTable();
            renderWorkerReport();
        } catch (e) {
            console.error('sectionWorkers init error', e);
        }
    }
    if (target === 'sectionSales') {
        try {
            renderSalesTable();
            updateSalesDropdown();
        } catch (e) {
            console.error('sectionSales init error', e);
        }
    }
    if (target === 'sectionTalaf') {
        try {
            renderTalafTable();
            updateTalafStats();
            updateTalafDropdown();
        } catch (e) {
            console.error('sectionTalaf init error', e);
        }
    }
    if (target === 'sectionInventory') {
        try {
            renderProfessionalInventory();
        } catch (e) {
            console.error('renderProfessionalInventory error', e);
        }
    }
    if (target === 'sectionReports') {
        try {
            renderWorkerReport();
        } catch (e) {
            console.error('renderWorkerReport error', e);
        }
    }
    if (target === 'sectionPayroll') {
        try {
            populatePayrollWorkerDropdowns();
            initPayrollDate();
            renderAttendanceTable();
            renderPayrollSummary();
            updatePayrollStats();
            renderQuickClockList();
        } catch (e) {
            console.error('sectionPayroll init error', e);
        }
    }
    if (target === 'sectionSettings') {
        try {
            renderCatalogTable();
        } catch (e) {
            console.error('sectionSettings init error', e);
        }
    }
});

// ==================== TAB SWITCHING (Excel Paste) ====================
function switchTab(tab) {
    document.getElementById('tabContentManual').style.display = tab === 'manual' ? 'block' : 'none';
    document.getElementById('tabContentPaste').style.display = tab === 'paste' ? 'block' : 'none';
    document.getElementById('tabManual').classList.toggle('active', tab === 'manual');
    document.getElementById('tabPaste').classList.toggle('active', tab === 'paste');
    if (tab === 'paste') {
        pasteArea.value = '';
        pastePreview.innerHTML = '';
        setTimeout(() => pasteArea.focus(), 100);
    }
}

// ==================== DATA PERSISTENCE ====================
function loadData(key, defaultValue) {
    const saved = localStorage.getItem(key);
    return saved ? JSON.parse(saved) : defaultValue;
}

function saveData(key, data) {
    localStorage.setItem(key, JSON.stringify(data));
}

let products = loadData('sticker_products', [
    { id: 1, batchNo: '230171', code: '01-MR0063', name: 'Glinmet 50mg/1000mg F.C Tab', qty: 13131, location: 'G3', stickerQty: 14399, price: 34000, pNumber: '23/3308_0002', company: 'ERA MEDICAL', receivedDate: '2023-07-31', expDate: '2026-02-01' }
]);

let workers = loadData('sticker_workers', [
    { id: 1, name: 'احمد علی', phone: '07501234567', job: 'کڕێکار', salary: 450000, joinDate: '2024-01-10', address: 'هەولێر - تەیراوە' }
]);

let sales = loadData('sticker_sales', []);
let talaf = loadData('sticker_talaf', []);
let catalog = loadData('sticker_catalog', []);
let activeSelectedShelf = null; // Declare early to avoid temporal dead zone crashes on startup
let attendance = loadData('sticker_attendance', []); // Declare early to avoid TDZ crashes on startup
let clockInLogs = loadData('sticker_clock_in_logs', {}); // Declare early to avoid TDZ crashes on startup

// Apply dark theme immediately if saved
if (localStorage.getItem('theme') === 'dark') {
    document.body.classList.add('dark-theme');
}

// --- Data Sanitization / Migration ---
// Ensure products is an array, has no null elements, and all items have a unique ID and default values
let productsUpdated = false;
if (!Array.isArray(products)) {
    products = [];
    productsUpdated = true;
} else {
    products = products.filter(p => p !== null && p !== undefined);
    products.forEach((p, idx) => {
        if (!p.id) {
            p.id = Date.now() + idx + Math.floor(Math.random() * 1000);
            productsUpdated = true;
        }
        // Sanitize properties to prevent runtime TypeErrors (e.g. on toLocaleString)
        if (p.qty === undefined || p.qty === null || isNaN(p.qty)) {
            p.qty = 0;
            productsUpdated = true;
        } else if (typeof p.qty === 'string') {
            p.qty = parseInt(p.qty.replace(/,/g, '')) || 0;
            productsUpdated = true;
        }
        
        if (p.stickerQty === undefined || p.stickerQty === null || isNaN(p.stickerQty)) {
            p.stickerQty = 0;
            productsUpdated = true;
        } else if (typeof p.stickerQty === 'string') {
            p.stickerQty = parseInt(p.stickerQty.replace(/,/g, '')) || 0;
            productsUpdated = true;
        }

        if (p.price === undefined || p.price === null || isNaN(p.price)) {
            p.price = 0;
            productsUpdated = true;
        } else if (typeof p.price === 'string') {
            p.price = parseInt(p.price.replace(/,/g, '')) || 0;
            productsUpdated = true;
        }

        if (!p.name) { p.name = 'بێ ناو'; productsUpdated = true; }
        if (!p.batchNo) { p.batchNo = '-'; productsUpdated = true; }
        if (!p.code) { p.code = '-'; productsUpdated = true; }
        if (!p.location) { p.location = '-'; productsUpdated = true; }
        if (!p.pNumber) { p.pNumber = '-'; productsUpdated = true; }
        if (!p.company) { p.company = '-'; productsUpdated = true; }
    });
}
if (productsUpdated) {
    saveData('sticker_products', products);
}

// Ensure workers is an array, has no null elements, and all items have a unique ID and default values
let workersUpdated = false;
if (!Array.isArray(workers)) {
    workers = [];
    workersUpdated = true;
} else {
    workers = workers.filter(w => w !== null && w !== undefined);
    workers.forEach((w, idx) => {
        if (!w.id) {
            w.id = Date.now() + 10000 + idx + Math.floor(Math.random() * 1000);
            workersUpdated = true;
        }
        if (!w.name) { w.name = 'کڕێکاری بێ ناو'; workersUpdated = true; }
    });
}
if (workersUpdated) {
    saveData('sticker_workers', workers);
}

// ==================== STICKER LOGIC ====================
let isStickerQtyManuallyEdited = false;

const closeModal = () => {
    stickerModal.classList.remove('active');
    stickerForm.reset();
    pasteArea.value = '';
    pastePreview.innerHTML = '';
    switchTab('manual');
    isStickerQtyManuallyEdited = false;
    if (window.stopActiveCameraScanner) window.stopActiveCameraScanner();
};

function populateAutofillDatalist() {
    const datalist = document.getElementById('existingProductsList');
    if (!datalist) return;
    const uniqueNames = [...new Set(products.map(p => p.name))].filter(Boolean);
    datalist.innerHTML = '';
    uniqueNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        datalist.appendChild(option);
    });
}

function findProductReference(field, value) {
    if (!value) return null;
    const valLower = value.trim().toLowerCase();
    // 1. Search in catalog
    let match = catalog.find(item => item[field] && String(item[field]).trim().toLowerCase() === valLower);
    if (match) return match;
    // 2. Search in active products
    match = products.find(item => item[field] && String(item[field]).trim().toLowerCase() === valLower);
    if (match) return match;
    return null;
}

const batchNoInput = document.getElementById('batchNo');
if (batchNoInput) {
    batchNoInput.addEventListener('input', () => {
        const enteredVal = batchNoInput.value.trim();
        if (!enteredVal) return;
        const match = findProductReference('batchNo', enteredVal);
        if (match) {
            if (match.code) document.getElementById('productCode').value = match.code;
            if (match.name) document.getElementById('productName').value = match.name;
            if (match.price) document.getElementById('price').value = match.price;
            if (match.location) document.getElementById('location').value = match.location;
            if (match.company) document.getElementById('company').value = match.company;
            if (match.pNumber) document.getElementById('pNumber').value = match.pNumber;
            if (match.expDate) document.getElementById('expDate').value = match.expDate;
            showToast('✨ زانیارییەکان بەپێی باچ خۆکار پڕکرانەوە!');
        }
    });
}

const productNameInput = document.getElementById('productName');
if (productNameInput) {
    productNameInput.addEventListener('input', () => {
        const enteredVal = productNameInput.value.trim();
        if (!enteredVal) return;
        const match = findProductReference('name', enteredVal);
        if (match) {
            if (match.batchNo) document.getElementById('batchNo').value = match.batchNo;
            if (match.code) document.getElementById('productCode').value = match.code;
            if (match.price) document.getElementById('price').value = match.price;
            if (match.location) document.getElementById('location').value = match.location;
            if (match.company) document.getElementById('company').value = match.company;
            if (match.pNumber) document.getElementById('pNumber').value = match.pNumber;
            if (match.expDate) document.getElementById('expDate').value = match.expDate;
            showToast('✨ زانیارییەکان بەپێی ناو خۆکار پڕکرانەوە!');
        }
    });
}

const productCodeInput = document.getElementById('productCode');
if (productCodeInput) {
    productCodeInput.addEventListener('input', () => {
        const enteredVal = productCodeInput.value.trim();
        if (!enteredVal) return;
        const match = findProductReference('code', enteredVal);
        if (match) {
            if (match.batchNo) document.getElementById('batchNo').value = match.batchNo;
            if (match.name) document.getElementById('productName').value = match.name;
            if (match.price) document.getElementById('price').value = match.price;
            if (match.location) document.getElementById('location').value = match.location;
            if (match.company) document.getElementById('company').value = match.company;
            if (match.pNumber) document.getElementById('pNumber').value = match.pNumber;
            if (match.expDate) document.getElementById('expDate').value = match.expDate;
            showToast('✨ زانیارییەکان بەپێی کۆد خۆکار پڕکرانەوە!');
        }
    });
}

const qtyInput = document.getElementById('qty');
const stickerQtyInput = document.getElementById('stickerQty');
if (stickerQtyInput) {
    stickerQtyInput.addEventListener('input', () => {
        isStickerQtyManuallyEdited = true;
    });
}
if (qtyInput && stickerQtyInput) {
    qtyInput.addEventListener('input', () => {
        if (!isStickerQtyManuallyEdited) {
            stickerQtyInput.value = qtyInput.value;
        }
    });
}

// Enter Key Navigation for stickerForm
const stickerFormInputs = [
    'batchNo',
    'productCode',
    'productName',
    'qty',
    'location',
    'stickerQty',
    'price',
    'pNumber',
    'company',
    'receivedDate',
    'expDate'
].map(id => document.getElementById(id)).filter(Boolean);

stickerFormInputs.forEach((input, index) => {
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault(); // Prevent submitting the form
            const nextInput = stickerFormInputs[index + 1];
            if (nextInput) {
                nextInput.focus();
                if (nextInput.select) nextInput.select();
            } else {
                // If it is the last input (expDate), we can submit the form
                const submitBtn = stickerForm.querySelector('button[type="submit"]');
                if (submitBtn) submitBtn.click();
            }
        }
    });
});
let cameraStream = null;
window.stopActiveCameraScanner = () => {
    if (cameraStream) {
        cameraStream.getTracks().forEach(track => track.stop());
        cameraStream = null;
    }
    const webcamVideo = document.getElementById('webcamVideo');
    if (webcamVideo) webcamVideo.srcObject = null;
    const cameraScannerContainer = document.getElementById('cameraScannerContainer');
    if (cameraScannerContainer) cameraScannerContainer.style.display = 'none';
};

const startCameraBtn = document.getElementById('startCameraBtn');
if (startCameraBtn) {
    startCameraBtn.addEventListener('click', async () => {
        try {
            cameraStream = await navigator.mediaDevices.getUserMedia({
                video: { facingMode: 'environment', width: { ideal: 1280 }, height: { ideal: 720 } }
            });
            const webcamVideo = document.getElementById('webcamVideo');
            if (webcamVideo) {
                webcamVideo.srcObject = cameraStream;
                const container = document.getElementById('cameraScannerContainer');
                if (container) container.style.display = 'flex';
                showToast('🎥 کامێرا چالاک کرا. لایبڵەکە ڕووبەڕووی کامێراکە بگرە.');
            }
        } catch (err) {
            console.error("Camera access error: ", err);
            showToast('❌ نەتوانرا دەست بە کامێرا بگات. تکایە مۆڵەتی پێویست بدە بە بەرنامەکە.');
        }
    });
}

const closeCameraBtn = document.getElementById('closeCameraBtn');
if (closeCameraBtn) {
    closeCameraBtn.addEventListener('click', () => {
        window.stopActiveCameraScanner();
    });
}

const captureBtn = document.getElementById('captureBtn');
if (captureBtn) {
    captureBtn.addEventListener('click', () => {
        const webcamVideo = document.getElementById('webcamVideo');
        const webcamCanvas = document.getElementById('webcamCanvas');
        if (!cameraStream || !webcamVideo || !webcamCanvas) return;

        const width = webcamVideo.videoWidth;
        const height = webcamVideo.videoHeight;
        webcamCanvas.width = width;
        webcamCanvas.height = height;

        const ctx = webcamCanvas.getContext('2d');
        ctx.drawImage(webcamVideo, 0, 0, width, height);

        webcamCanvas.toBlob((blob) => {
            window.stopActiveCameraScanner();
            showToast('⏳ خەریکە وێنەکە لە کامێراکەوە دەخوێنێتەوە، تکایە چاوەڕێبە...');

            Tesseract.recognize(
                blob,
                'eng',
                { logger: m => console.log(m) }
            ).then(({ data: { text } }) => {
                console.log("Webcam OCR Text extracted: ", text);
                
                const nameMatch = text.match(/Product\s*Name\s*[:\-]?\s*(.+)/i);
                const importerMatch = text.match(/Importer\s*[:\-]?\s*(.+)/i);
                const pNoMatch = text.match(/P\.?NO\s*[:\-]?\s*([A-Z0-9\-]+)/i);
                const qtyMatch = text.match(/Quantity\s*[:\-]?\s*(\d+)/i);
                const shipDateMatch = text.match(/Ship\s*Date\s*[:\-]?\s*([\d\-]+)/i);
                
                let autoFilledCount = 0;

                if (nameMatch) {
                    document.getElementById('productName').value = nameMatch[1].trim();
                    autoFilledCount++;
                }
                if (importerMatch) {
                    document.getElementById('company').value = importerMatch[1].trim();
                    autoFilledCount++;
                }
                if (pNoMatch) {
                    document.getElementById('batchNo').value = pNoMatch[1].trim();
                    autoFilledCount++;
                }
                if (qtyMatch) {
                    document.getElementById('stickerQty').value = parseInt(qtyMatch[1]);
                    autoFilledCount++;
                }
                
                if (shipDateMatch) {
                    const parts = shipDateMatch[1].split('-');
                    if (parts.length === 3) {
                        let year = parts[2].trim();
                        if (year.length === 3 && year.startsWith('0')) year = '20' + year.substring(1);
                        else if (year.length === 2) year = '20' + year;
                        document.getElementById('receivedDate').value = `${year}-${parts[1].trim()}-${parts[0].trim()}`;
                        autoFilledCount++;
                    }
                }

                const reqMatch = text.match(/(REQ-[A-Z0-9\-]+\s+[0-9\-]+)/i);
                if (reqMatch) {
                    document.getElementById('productCode').value = reqMatch[1].trim();
                    autoFilledCount++;
                }

                const permitMatch = text.match(/(\d{2}-\d{4,})/);
                if (permitMatch) {
                    document.getElementById('pNumber').value = permitMatch[1].trim();
                    autoFilledCount++;
                }

                if (autoFilledCount > 0) {
                    showToast(`✅ ${autoFilledCount} خانە بە سەرکەوتوویی لە کامێراکەوە پڕکرانەوە!`);
                } else {
                    showToast('⚠️ نەتوانرا هیچ زانیارییەک لە کامێراکەوە بخوێنرێتەوە. تکایە دووبارە تاقی بکەرەوە.');
                }
            }).catch(err => {
                console.error("Webcam OCR Error: ", err);
                showToast('❌ هەڵەیەک ڕوویدا لە کاتی خوێندنەوەی وێنەی کامێراکەدا.');
            });
        }, 'image/jpeg');
    });
}

const labelImageInput = document.getElementById('labelImageInput');
if (labelImageInput) {
    labelImageInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        showToast('⏳ خەریکە وێنەکە دەخوێنێتەوە، تکایە چاوەڕێبە...');

        Tesseract.recognize(
            file,
            'eng',
            { logger: m => console.log(m) }
        ).then(({ data: { text } }) => {
            console.log("OCR Text extracted: ", text);
            
            const nameMatch = text.match(/Product\s*Name\s*[:\-]?\s*(.+)/i);
            const importerMatch = text.match(/Importer\s*[:\-]?\s*(.+)/i);
            const pNoMatch = text.match(/P\.?NO\s*[:\-]?\s*([A-Z0-9\-]+)/i);
            const qtyMatch = text.match(/Quantity\s*[:\-]?\s*(\d+)/i);
            const shipDateMatch = text.match(/Ship\s*Date\s*[:\-]?\s*([\d\-]+)/i);
            
            let autoFilledCount = 0;

            if (nameMatch) {
                document.getElementById('productName').value = nameMatch[1].trim();
                autoFilledCount++;
            }
            if (importerMatch) {
                document.getElementById('company').value = importerMatch[1].trim();
                autoFilledCount++;
            }
            if (pNoMatch) {
                document.getElementById('batchNo').value = pNoMatch[1].trim();
                autoFilledCount++;
            }
            if (qtyMatch) {
                document.getElementById('stickerQty').value = parseInt(qtyMatch[1]);
                autoFilledCount++;
            }
            
            if (shipDateMatch) {
                const parts = shipDateMatch[1].split('-');
                if (parts.length === 3) {
                    let year = parts[2].trim();
                    if (year.length === 3 && year.startsWith('0')) year = '20' + year.substring(1);
                    else if (year.length === 2) year = '20' + year;
                    document.getElementById('receivedDate').value = `${year}-${parts[1].trim()}-${parts[0].trim()}`;
                    autoFilledCount++;
                }
            }

            const reqMatch = text.match(/(REQ-[A-Z0-9\-]+\s+[0-9\-]+)/i);
            if (reqMatch) {
                document.getElementById('productCode').value = reqMatch[1].trim();
                autoFilledCount++;
            }

            const permitMatch = text.match(/(\d{2}-\d{4,})/);
            if (permitMatch) {
                document.getElementById('pNumber').value = permitMatch[1].trim();
                autoFilledCount++;
            }

            if (autoFilledCount > 0) {
                showToast(`✅ ${autoFilledCount} خانە بە سەرکەوتوویی لە وێنەکەوە پڕکرانەوە!`);
            } else {
                showToast('⚠️ نەتوانرا هیچ زانیارییەک بخوێنرێتەوە. تکایە دەستی داخڵی بکە یان وێنەکە ڕوونتر بکە.');
            }
            
            labelImageInput.value = '';
        }).catch(err => {
            console.error("OCR Error: ", err);
            showToast('❌ هەڵەیەک ڕوویدا لە کاتی خوێندنەوەی وێنەکەدا.');
            labelImageInput.value = '';
        });
    });
}

openModalBtn.addEventListener('click', () => {
    stickerModal.classList.add('active');
    populateAutofillDatalist();

    // Set default received date to today
    const receivedDateInput = document.getElementById('receivedDate');
    if (receivedDateInput) {
        receivedDateInput.value = new Date().toISOString().split('T')[0];
    }

    // Autofocus the batchNo field
    setTimeout(() => {
        const batchNoInput = document.getElementById('batchNo');
        if (batchNoInput) {
            batchNoInput.focus();
            if (batchNoInput.select) batchNoInput.select();
        }
    }, 100);
});
closeModalBtn.addEventListener('click', closeModal);

stickerForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const newProduct = {
        id: Date.now(),
        batchNo: document.getElementById('batchNo').value,
        code: document.getElementById('productCode').value,
        name: document.getElementById('productName').value,
        qty: parseInt(document.getElementById('qty').value) || 0,
        location: document.getElementById('location').value,
        stickerQty: parseInt(document.getElementById('stickerQty').value) || 0,
        price: parseInt(document.getElementById('price').value) || 0,
        pNumber: document.getElementById('pNumber').value,
        company: document.getElementById('company').value,
        receivedDate: document.getElementById('receivedDate').value,
        expDate: document.getElementById('expDate').value
    };
    products.push(newProduct);
    saveData('sticker_products', products);

    // Auto-save/update reference catalog
    if (newProduct.name) {
        const existingIndex = catalog.findIndex(item => 
            (newProduct.batchNo && item.batchNo && item.batchNo.toLowerCase() === newProduct.batchNo.toLowerCase()) || 
            (newProduct.code && item.code && item.code.toLowerCase() === newProduct.code.toLowerCase())
        );
        const templateData = {
            id: existingIndex >= 0 ? catalog[existingIndex].id : Date.now() + Math.random(),
            batchNo: newProduct.batchNo,
            code: newProduct.code,
            name: newProduct.name,
            price: newProduct.price,
            location: newProduct.location,
            company: newProduct.company,
            pNumber: newProduct.pNumber,
            expDate: newProduct.expDate
        };
        if (existingIndex >= 0) {
            catalog[existingIndex] = templateData;
        } else {
            catalog.push(templateData);
        }
        saveData('sticker_catalog', catalog);
    }

    renderTable();
    updateSalesDropdown();
    closeModal();
    showToast('بەرهەمەکە بە سەرکەوتوویی زیادکرا! ✅');
});

function renderTable() {
    tableBody.innerHTML = '';
    products.forEach(p => {
        const exp = new Date(p.expDate);
        const now = new Date();
        const diffMonths = (exp - now) / (1000 * 60 * 60 * 24 * 30);
        let expClass = diffMonths < 0 ? 'low-stock' : (diffMonths < 6 ? 'exp-soon' : 'in-stock');

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><span class="batch-badge">${p.batchNo}</span></td>
            <td><code class="code-cell">${p.code}</code></td>
            <td class="product-name-cell">${p.name}</td>
            <td>${p.qty.toLocaleString()}</td>
            <td><span class="location-badge">${p.location}</span></td>
            <td><span class="stock-badge in-stock">${p.stickerQty.toLocaleString()}</span></td>
            <td>${p.price.toLocaleString()}</td>
            <td>${p.pNumber}</td>
            <td>${p.company}</td>
            <td>${formatDate(p.receivedDate)}</td>
            <td><span class="stock-badge ${expClass}">${formatDate(p.expDate)}</span></td>
            <td>
                <div class="action-btns">
                    <button class="action-btn sell" title="فرۆشتن" onclick="openQuickSell(${p.id})">🛒</button>
                    <button class="action-btn edit" onclick="editProduct(${p.id})">✏️</button>
                    <button class="action-btn delete" onclick="deleteProduct(${p.id})">🗑️</button>
                </div>
            </td>
        `;
        tableBody.appendChild(tr);
    });
    updateStats();
    renderProfessionalInventory();
}

function updateStats() {
    document.getElementById('totalProducts').innerText = products.length;

    const totalVal = products.reduce((acc, p) => acc + (p.price * p.qty), 0);
    document.getElementById('totalValue').innerText = totalVal.toLocaleString() + ' IQD';

    const now = new Date();
    document.getElementById('lowStickerStock').innerText = products.filter(p => {
        if (!p.expDate) return false;
        const exp = new Date(p.expDate);
        const diffMonths = (exp - now) / (1000 * 60 * 60 * 24 * 30);
        return diffMonths < 6;
    }).length;

    // Calculate Total Sales Today
    const today = new Date().toDateString();
    const salesToday = sales.filter(s => new Date(s.date).toDateString() === today);
    const totalStickersSoldToday = salesToday.reduce((acc, s) => acc + (s.stickerQty || 0), 0);
    document.getElementById('totalSalesToday').innerText = totalStickersSoldToday.toLocaleString();

    // 1. DYNAMIC STICKER USAGE METRICS
    const totalStickersInStock = products.reduce((acc, p) => acc + (p.stickerQty || 0), 0);
    const totalStickersSoldAllTime = sales.reduce((acc, s) => acc + (s.stickerQty || 0), 0);
    const totalProcessed = totalStickersInStock + totalStickersSoldAllTime;
    const remainingPercent = totalProcessed > 0 ? Math.round((totalStickersInStock / totalProcessed) * 100) : 100;

    const usePercentEl = document.getElementById('stickerUsagePercent');
    const useProgressEl = document.getElementById('stickerUsageProgressBar');
    const soldTodayEl = document.getElementById('stickersSoldTodaySpan');
    const totalStockEl = document.getElementById('stickersTotalStockSpan');

    if (usePercentEl) usePercentEl.innerText = remainingPercent + '%';
    if (useProgressEl) useProgressEl.style.width = remainingPercent + '%';
    if (soldTodayEl) soldTodayEl.innerText = 'فرۆشراوی ئەمڕۆ: ' + totalStickersSoldToday.toLocaleString();
    if (totalStockEl) totalStockEl.innerText = 'ماوە لە کۆگا: ' + totalStickersInStock.toLocaleString();

    // 2. DYNAMIC WORKER LEADERBOARD
    const leaderboardEl = document.getElementById('workerLeaderboardList');
    if (leaderboardEl) {
        // Aggregate stickers per worker
        const wSales = {};
        sales.forEach(sale => {
            if (sale.assignedWorkers) {
                sale.assignedWorkers.forEach(w => {
                    if (!wSales[w.id]) {
                        wSales[w.id] = { id: w.id, name: w.name, total: 0 };
                    }
                    wSales[w.id].total += (w.qty || 0);
                });
            }
        });

        // Convert to sorted array
        const sortedWorkers = Object.values(wSales).sort((a, b) => b.total - a.total);

        if (sortedWorkers.length === 0) {
            leaderboardEl.innerHTML = '<div style="text-align: center; color: #64748b; padding: 20px 0; font-size: 0.85rem;">هیچ چالاکییەکی تۆمارکراو نییە.</div>';
        } else {
            leaderboardEl.innerHTML = '';
            sortedWorkers.slice(0, 5).forEach((w, index) => {
                let medal = '';
                if (index === 0) medal = '🥇';
                else if (index === 1) medal = '🥈';
                else if (index === 2) medal = '🥉';
                else medal = `<span style="font-weight:700; color:#475569; width:24px; text-align:center; display:inline-block;">#${index + 1}</span>`;

                const workerObj = workers.find(wr => wr.id == w.id) || { job: 'کڕێکار' };
                const jobTitle = workerObj.job || 'کڕێکار';

                const row = document.createElement('div');
                row.style.cssText = 'display: flex; align-items: center; justify-content: space-between; padding: 10px 14px; background: rgba(255, 255, 255, 0.02); border-radius: 12px; border: 1px solid rgba(255, 255, 255, 0.04); transition: transform 0.2s;';
                row.innerHTML = `
                    <div style="display: flex; align-items: center; gap: 12px;">
                        <span style="font-size: 1.3rem; display: flex; align-items: center; justify-content: center; width: 24px;">${medal}</span>
                        <div>
                            <div style="font-weight: 600; color: #1f2937; font-size: 0.9rem;">${w.name}</div>
                            <div style="font-size: 0.72rem; color: #64748b;">${jobTitle}</div>
                        </div>
                    </div>
                    <div style="text-align: right;">
                        <strong style="color: #10b981; font-size: 0.95rem;">${w.total.toLocaleString()}</strong>
                        <span style="font-size: 0.72rem; color: #64748b; display: block;">ستیکەر</span>
                    </div>
                `;
                leaderboardEl.appendChild(row);
            });
        }
    }
    if (window.updateNotifications) window.updateNotifications();
}

function deleteProduct(id) {
    if (confirm('دڵنیایت؟')) {
        products = products.filter(p => p.id !== id);
        saveData('sticker_products', products);
        renderTable();
        showToast('بەرهەمەکە سڕایەوە! 🗑️');
    }
}

function editProduct(id) {
    const p = products.find(pr => pr.id === id);
    if (!p) return;
    document.getElementById('batchNo').value = p.batchNo;
    document.getElementById('productCode').value = p.code;
    document.getElementById('productName').value = p.name;
    document.getElementById('qty').value = p.qty;
    document.getElementById('location').value = p.location;
    document.getElementById('stickerQty').value = p.stickerQty;
    document.getElementById('price').value = p.price;
    document.getElementById('pNumber').value = p.pNumber;
    document.getElementById('company').value = p.company;
    document.getElementById('receivedDate').value = p.receivedDate;
    document.getElementById('expDate').value = p.expDate;
    products = products.filter(pr => pr.id !== id);
    stickerModal.classList.add('active');

    // Autofocus the batchNo field
    setTimeout(() => {
        const batchNoInput = document.getElementById('batchNo');
        if (batchNoInput) {
            batchNoInput.focus();
            if (batchNoInput.select) batchNoInput.select();
        }
    }, 100);
}

// ==================== WORKER LOGIC ====================
const closeWorkerModal = () => {
    workerModal.classList.remove('active');
    workerForm.reset();
};

openWorkerModalBtn.addEventListener('click', () => workerModal.classList.add('active'));
closeWorkerModalBtn.addEventListener('click', closeWorkerModal);
cancelWorkerBtn.addEventListener('click', closeWorkerModal);

workerForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const newWorker = {
        id: Date.now(),
        name: document.getElementById('workerName').value,
        phone: document.getElementById('workerPhone').value,
        job: document.getElementById('workerJob').value,
        salary: parseInt(document.getElementById('workerSalary').value),
        joinDate: document.getElementById('workerJoinDate').value,
        address: document.getElementById('workerAddress').value
    };
    workers.push(newWorker);
    saveData('sticker_workers', workers);
    renderWorkerTable();
    updateStats();
    closeWorkerModal();
    showToast('کڕێکارەکە بە سەرکەوتوویی زیادکرا! 👷');
});

function renderWorkerTable() {
    workerTableBody.innerHTML = '';
    workers.forEach(w => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${w.name}</td>
            <td>${w.phone}</td>
            <td>${w.job}</td>
            <td>${w.salary.toLocaleString()}</td>
            <td>${formatDate(w.joinDate)}</td>
            <td>${w.address}</td>
            <td>
                <div class="action-btns">
                    <button class="action-btn edit" onclick="editWorker(${w.id})">✏️</button>
                    <button class="action-btn delete" onclick="deleteWorker(${w.id})">🗑️</button>
                </div>
            </td>
        `;
        workerTableBody.appendChild(tr);
    });
    document.getElementById('totalWorkers').innerText = workers.length;
    if (window.renderQuickClockList) window.renderQuickClockList();
}

function deleteWorker(id) {
    if (confirm('دڵنیایت لە سڕینەوەی ئەم کڕێکارە؟')) {
        workers = workers.filter(w => w.id !== id);
        saveData('sticker_workers', workers);
        renderWorkerTable();
        updateStats();
        showToast('کڕێکارەکە سڕایەوە! 🗑️');
    }
}

function editWorker(id) {
    const w = workers.find(wr => wr.id === id);
    if (!w) return;
    document.getElementById('workerName').value = w.name;
    document.getElementById('workerPhone').value = w.phone;
    document.getElementById('workerJob').value = w.job;
    document.getElementById('workerSalary').value = w.salary;
    document.getElementById('workerJoinDate').value = w.joinDate;
    document.getElementById('workerAddress').value = w.address;
    workers = workers.filter(wr => wr.id !== id);
    workerModal.classList.add('active');
}

// ==================== SALES LOGIC ====================
function addWorkerRow() {
    const row = document.createElement('div');
    row.className = 'worker-assignment-row';

    let workerOptions = workers.map(w => `<option value="${w.id}">${w.name}</option>`).join('');

    row.innerHTML = `
        <div class="input-group">
            <label style="font-size: 0.7rem;">کڕێکار</label>
            <select class="worker-select" required>
                <option value="">هەڵبژێرە...</option>
                ${workerOptions}
            </select>
        </div>
        <div class="input-group">
            <label style="font-size: 0.7rem;">بڕی ستیکەر</label>
            <input type="number" class="worker-qty-input" placeholder="0" required>
        </div>
        <button type="button" class="remove-worker-btn" onclick="this.parentElement.remove()">✕</button>
    `;
    workerAssignmentList.appendChild(row);
}

addWorkerToSaleBtn.addEventListener('click', addWorkerRow);

function updateSalesDropdown() {
    if (!salesProductSelect) return;
    salesProductSelect.innerHTML = '<option value="">هەڵبژێرە...</option>';
    if (!Array.isArray(products)) return;
    products.forEach(p => {
        if (!p) return;
        const option = document.createElement('option');
        option.value = p.id || '';
        const name = p.name || 'بێ ناو';
        const batch = p.batchNo || '-';
        const location = p.location || '-';
        option.textContent = `${name} (باچ: ${batch} | 📍 شوێن: ${location})`;
        salesProductSelect.appendChild(option);
    });
}

salesProductSelect.addEventListener('change', () => {
    const productId = salesProductSelect.value;
    const indicator = document.getElementById('stockIndicator');
    const stockVal = document.getElementById('currentStockVal');
    const stickerVal = document.getElementById('currentStickerStockVal');
    const locationVal = document.getElementById('currentLocationVal');

    if (!productId || productId === 'undefined') {
        indicator.style.display = 'none';
        return;
    }

    const p = products.find(prod => prod && prod.id && prod.id.toString() === productId.toString());
    if (p) {
        stockVal.textContent = (p.qty || 0).toLocaleString();
        stickerVal.textContent = (p.stickerQty || 0).toLocaleString();
        if (locationVal) locationVal.textContent = p.location || '-';
        indicator.style.display = 'flex';
    } else {
        indicator.style.display = 'none';
    }
});

addToCartForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const productId = document.getElementById('salesProductSelect').value;
    const sQty = parseInt(document.getElementById('salesQty').value);

    // Collect Worker Data
    const workerRows = document.querySelectorAll('.worker-assignment-row');
    const assignedWorkers = [];
    let totalStickerQty = 0;

    workerRows.forEach(row => {
        const wSelect = row.querySelector('.worker-select');
        const wId = wSelect.value;
        const wName = wSelect.options[wSelect.selectedIndex].text;
        const wQty = parseInt(row.querySelector('.worker-qty-input').value) || 0;
        assignedWorkers.push({ id: wId, name: wName, qty: wQty });
        totalStickerQty += wQty;
    });

    if (assignedWorkers.length === 0) {
        showToast('⚠️ تکایە لانیکەم کڕێکارێک دیاری بکە!');
        return;
    }

    const productIndex = products.findIndex(p => p && p.id && p.id.toString() === productId.toString());
    if (productIndex === -1) {
        showToast('⚠️ هەڵەیەک لە بەرهەمی هەڵبژێردراو هەیە!');
        return;
    }

    // Calculate total already in cart for this product
    const alreadyInCartQty = currentInvoiceCart.filter(item => item.productId.toString() === productId.toString()).reduce((sum, item) => sum + item.sQty, 0);
    const alreadyInCartStickers = currentInvoiceCart.filter(item => item.productId.toString() === productId.toString()).reduce((sum, item) => sum + item.totalStickerQty, 0);

    if (products[productIndex].qty < (sQty + alreadyInCartQty) || products[productIndex].stickerQty < (totalStickerQty + alreadyInCartStickers)) {
        showToast('⚠️ بڕی بەردەست بەش ناکات یان زۆرت خستۆتە ناو قائیمە!');
        return;
    }

    // Add to Cart
    currentInvoiceCart.push({
        id: Date.now() + Math.random(),
        productId: productId,
        productName: products[productIndex].name,
        sQty: sQty,
        totalStickerQty: totalStickerQty,
        assignedWorkers: assignedWorkers
    });

    addToCartForm.reset();
    workerAssignmentList.innerHTML = '';
    renderInvoiceCart();
    showToast('🛒 مادەکە خرایە ناو قائیمە!');
});

function renderInvoiceCart() {
    const section = document.getElementById('invoiceCartSection');
    const tbody = document.getElementById('invoiceCartTableBody');
    const count = document.getElementById('invoiceCartCount');
    const submitBtn = document.getElementById('submitInvoiceBtn');

    if (currentInvoiceCart.length === 0) {
        section.style.display = 'none';
        submitBtn.style.display = 'none';
        return;
    }

    section.style.display = 'block';
    submitBtn.style.display = 'block';
    count.innerText = currentInvoiceCart.length;
    tbody.innerHTML = '';

    currentInvoiceCart.forEach((item, index) => {
        const tr = document.createElement('tr');
        const workersHtml = item.assignedWorkers.map(w => `<span class="worker-tag">${w.name}: ${w.qty}</span>`).join(' ');
        tr.innerHTML = `
            <td style="font-weight: 600;">${item.productName}</td>
            <td><span class="stock-badge in-stock">${item.sQty.toLocaleString()}</span></td>
            <td><span class="stock-badge blue">${item.totalStickerQty.toLocaleString()}</span></td>
            <td>${workersHtml}</td>
            <td style="text-align: center;">
                <button type="button" class="action-btn delete" onclick="removeCartItem(${index})">🗑️</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

window.removeCartItem = function(index) {
    currentInvoiceCart.splice(index, 1);
    renderInvoiceCart();
};

document.getElementById('submitInvoiceBtn').addEventListener('click', () => {
    const destination = document.getElementById('salesDestination').value.trim();
    if (!destination) {
        showToast('⚠️ تکایە شوێنی مەبەست (Destination) بنووسە!');
        document.getElementById('salesDestination').focus();
        return;
    }

    if (currentInvoiceCart.length === 0) return;

    // Process all cart items
    currentInvoiceCart.forEach(item => {
        const pIndex = products.findIndex(p => p && p.id && p.id.toString() === item.productId.toString());
        if (pIndex !== -1) {
            products[pIndex].qty -= item.sQty;
            products[pIndex].stickerQty -= item.totalStickerQty;

            // Record Sale
            sales.unshift({
                id: Date.now() + Math.random(),
                productName: item.productName,
                destination: destination,
                qty: item.sQty,
                stickerQty: item.totalStickerQty,
                assignedWorkers: item.assignedWorkers,
                date: new Date().toISOString()
            });
        }
    });

    saveData('sticker_products', products);
    saveData('sticker_sales', sales);

    renderTable();
    renderSalesTable();
    updateSalesDropdown();
    
    currentInvoiceCart = [];
    renderInvoiceCart();
    document.getElementById('salesDestination').value = '';
    
    // Refresh indicator if needed
    const productId = document.getElementById('salesProductSelect').value;
    if (productId) {
        document.getElementById('salesProductSelect').dispatchEvent(new Event('change'));
    }

    showToast('✅ تەواوی قائیمەکە سەرکەوتووانە تۆمارکرا!');
});

function renderSalesTable() {
    salesTableBody.innerHTML = '';
    sales.forEach(s => {
        const workersHtml = (s.assignedWorkers || []).map(w => `<span class="worker-tag">${w.name}: ${w.qty}</span>`).join(' ');

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${s.productName}</td>
            <td><span class="location-badge">${s.destination || '-'}</span></td>
            <td style="color: #ef4444;">-${s.qty.toLocaleString()}</td>
            <td style="color: #ef4444;">-${s.stickerQty.toLocaleString()}</td>
            <td>${workersHtml}</td>
            <td>${new Date(s.date).toLocaleString('ku-IQ')}</td>
            <td style="text-align:center;">
                <div class="action-btns" style="justify-content:center;">
                    <button class="action-btn edit" title="دەستکاریکردن" onclick="editSale('${s.id}')">✏️</button>
                    <button class="action-btn delete" title="سڕینەوە" onclick="deleteSale('${s.id}')">🗑️</button>
                </div>
            </td>
        `;
        salesTableBody.appendChild(tr);
    });
}

// ==================== DELETE SALE ====================
window.deleteSale = function(id) {
    const idx = sales.findIndex(s => String(s.id) === String(id));
    if (idx === -1) return;

    if (!confirm('دڵنیایت لە سڕینەوەی ئەم تۆمارەی فرۆشتنە؟\nبڕەکانی دەگەڕێنێتەوە بۆ کۆگا.')) return;

    const s = sales[idx];

    // Restore stock to the product
    const pIdx = products.findIndex(p => p && p.name === s.productName);
    if (pIdx !== -1) {
        products[pIdx].qty = (products[pIdx].qty || 0) + (s.qty || 0);
        products[pIdx].stickerQty = (products[pIdx].stickerQty || 0) + (s.stickerQty || 0);
        saveData('sticker_products', products);
        renderTable();
        updateSalesDropdown();
    }

    sales.splice(idx, 1);
    saveData('sticker_sales', sales);
    renderSalesTable();
    updateStats();
    showToast('تۆمارەکە سڕایەوە و بڕەکانی گەڕاندرایەوە بۆ کۆگا! 🗑️');
};

// ==================== EDIT SALE ====================
window.editSale = function(id) {
    const s = sales.find(sale => String(sale.id) === String(id));
    if (!s) return;

    document.getElementById('editSaleId').value = s.id;
    document.getElementById('editSaleProductName').value = s.productName;
    document.getElementById('editSaleDestination').value = s.destination || '';
    document.getElementById('editSaleQty').value = s.qty;
    document.getElementById('editSaleStickerQty').value = s.stickerQty;

    const modal = document.getElementById('editSaleModal');
    modal.style.display = 'flex';
};

// Close edit sale modal
document.getElementById('closeEditSaleModalBtn').addEventListener('click', () => {
    document.getElementById('editSaleModal').style.display = 'none';
});
document.getElementById('cancelEditSaleBtn').addEventListener('click', () => {
    document.getElementById('editSaleModal').style.display = 'none';
});

// Save edited sale
document.getElementById('saveEditSaleBtn').addEventListener('click', () => {
    const id = document.getElementById('editSaleId').value;
    const newDestination = document.getElementById('editSaleDestination').value.trim();
    const newQty = parseInt(document.getElementById('editSaleQty').value) || 0;
    const newStickerQty = parseInt(document.getElementById('editSaleStickerQty').value) || 0;

    if (!newDestination) {
        showToast('⚠️ تکایە شوێنی مەبەست بنووسە!');
        return;
    }

    const idx = sales.findIndex(s => String(s.id) === String(id));
    if (idx === -1) { showToast('❌ تۆمارەکە نەدۆزرایەوە!'); return; }

    const oldSale = sales[idx];
    const qtyDiff = oldSale.qty - newQty;           // positive = we return stock, negative = we take more
    const stickerDiff = oldSale.stickerQty - newStickerQty;

    // Update product stock accordingly
    const pIdx = products.findIndex(p => p && p.name === oldSale.productName);
    if (pIdx !== -1) {
        const newProductQty = (products[pIdx].qty || 0) + qtyDiff;
        const newProductSticker = (products[pIdx].stickerQty || 0) + stickerDiff;

        if (newProductQty < 0 || newProductSticker < 0) {
            showToast('⚠️ بڕی نوێ زیاترە لە مەوجودی کۆگا!');
            return;
        }
        products[pIdx].qty = newProductQty;
        products[pIdx].stickerQty = newProductSticker;
        saveData('sticker_products', products);
        renderTable();
        updateSalesDropdown();
    }

    // Update sale record
    sales[idx] = {
        ...oldSale,
        destination: newDestination,
        qty: newQty,
        stickerQty: newStickerQty
    };

    saveData('sticker_sales', sales);
    renderSalesTable();
    updateStats();
    document.getElementById('editSaleModal').style.display = 'none';
    showToast('✅ تۆمارەکە بە سەرکەوتوویی گۆڕدرا!');
});



function openQuickSell(id) {
    const navItem = document.querySelector('.nav-item[data-target="sectionSales"]');
    if (navItem) navItem.click();
    setTimeout(() => {
        salesProductSelect.value = id;
        salesProductSelect.dispatchEvent(new Event('change'));
        document.getElementById('salesQty').focus();
    }, 100);
}

// ==================== INVENTORY LOGIC ====================
let currentInventoryView = 'table'; // 'table' or 'grid'

// Setup elements mapping to actual dashboard.html IDs
const invSearchInput = document.getElementById('filterName');
const invStatusFilter = document.getElementById('filterExpStatus');
const invSortFilter = document.getElementById('filterSortBy');
const invCompanyFilter = document.getElementById('filterCompany');
const invTableBody = document.getElementById('stickerTableBody');
const invGridView = document.getElementById('invGridView');
const invTableView = document.getElementById('invTableView');
const viewToggleTable = document.getElementById('viewToggleTable');
const viewToggleGrid = document.getElementById('viewToggleGrid');

// Global aliases to fix broken HTML bindings
window.renderTable = renderProfessionalInventory;
window.applyFilters = renderProfessionalInventory;
window.clearAllFilters = function() {
    if (document.getElementById('filterName')) document.getElementById('filterName').value = '';
    if (document.getElementById('filterBatch')) document.getElementById('filterBatch').value = '';
    if (document.getElementById('filterCode')) document.getElementById('filterCode').value = '';
    if (document.getElementById('filterLocation')) document.getElementById('filterLocation').value = '';
    if (document.getElementById('filterCompany')) document.getElementById('filterCompany').value = '';
    if (document.getElementById('filterExpStatus')) document.getElementById('filterExpStatus').value = '';
    if (document.getElementById('filterSortBy')) document.getElementById('filterSortBy').value = '';
    renderProfessionalInventory();
};
window.quickFilter = function(type) {
    if (document.getElementById('filterExpStatus')) document.getElementById('filterExpStatus').value = '';
    if (document.getElementById('filterCompany')) document.getElementById('filterCompany').value = '';
    if (type === 'expired') { document.getElementById('filterExpStatus').value = 'expired'; }
    if (type === 'soon') { document.getElementById('filterExpStatus').value = 'soon'; }
    if (type === 'low_stock') { /* Handle low stock */ }
    if (type === 'hawkary') { document.getElementById('filterCompany').value = 'HAWKARY'; }
    renderProfessionalInventory();
};

// Event Listeners for Filters
if (invSearchInput) invSearchInput.addEventListener('input', renderProfessionalInventory);
if (invStatusFilter) invStatusFilter.addEventListener('change', renderProfessionalInventory);
if (invSortFilter) invSortFilter.addEventListener('change', renderProfessionalInventory);
if (invCompanyFilter) invCompanyFilter.addEventListener('change', renderProfessionalInventory);

// Event Listeners for View Toggle
if (viewToggleTable) {
    viewToggleTable.addEventListener('click', () => {
        currentInventoryView = 'table';
        viewToggleTable.classList.add('active');
        viewToggleGrid.classList.remove('active');
        invTableView.style.display = 'block';
        invGridView.style.display = 'none';
        renderProfessionalInventory();
    });
}
if (viewToggleGrid) {
    viewToggleGrid.addEventListener('click', () => {
        currentInventoryView = 'grid';
        viewToggleGrid.classList.add('active');
        viewToggleTable.classList.remove('active');
        invGridView.style.display = 'grid';
        invTableView.style.display = 'none';
        renderProfessionalInventory();
    });
}

// Export Inventory to Excel
const inventoryExportBtn = document.getElementById('exportProductsBtn');
if (inventoryExportBtn) {
    inventoryExportBtn.addEventListener('click', () => {
        const filtered = getFilteredProducts();
        if (filtered.length === 0) {
            showToast('⚠️ لیستەکە خاڵیە و هیچ بەرهەمێک نییە بۆ هەناردەکردن!');
            return;
        }

        const wb = XLSX.utils.book_new();
        const dataRows = filtered.map(p => ({
            'ژمارەی باچ (Batch No)': p.batchNo,
            'کۆد (Code)': p.code,
            'ناوی بەرهەم (Product Name)': p.name,
            'بڕی کۆگا (Stock Qty)': p.qty,
            'بڕی ستیکەر (Stickers Qty)': p.stickerQty,
            'نرخی تاک (Unit Price)': p.price,
            'کۆی گشتی بەها (Total Value)': p.qty * p.price,
            'شوێن (Location)': p.location,
            'کۆمپانیا (Company)': p.company,
            'بەرواری وەرگرتن (Received Date)': formatDate(p.receivedDate),
            'بەرواری بەسەرچوون (Expiration Date)': formatDate(p.expDate)
        }));

        const ws = XLSX.utils.json_to_sheet(dataRows);
        XLSX.utils.book_append_sheet(wb, ws, 'مەوجودات');

        // RTL layout support for Arabic/Kurdish sheet
        ws['!dir'] = 'rtl';

        // Choose save destination
        if (electronDialog) {
            electronDialog.showSaveDialog({
                title: 'هەناردەکردنی مەوجودات بۆ نێو فایلی Excel',
                defaultPath: path.join(os.homedir(), 'Desktop', `Mewcudat_${Date.now()}.xlsx`),
                filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
            }).then(result => {
                if (!result.canceled && result.filePath) {
                    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
                    function s2ab(s) {
                        const buf = new ArrayBuffer(s.length);
                        const view = new Uint8Array(buf);
                        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
                        return buf;
                    }
                    fs.writeFileSync(result.filePath, Buffer.from(s2ab(wbout)));
                    showToast('📊 لیستەکە بە سەرکەوتوویی هەناردە کرا! ✅');
                }
            }).catch(err => {
                showToast('⚠️ کێشەیەک لە هەناردەکردن ڕوویدا!');
                console.error(err);
            });
        } else {
            // Fallback for non-electron testing/dev environments
            const filename = `Mewcudat_${Date.now()}.xlsx`;
            XLSX.writeFile(wb, filename);
            showToast(`📊 هەناردەکرا بۆ ${filename}! ✅`);
        }
    });
}

function updateShelfLiveView() {
    const shelfGrid = document.getElementById('shelfMapGrid');
    if (!shelfGrid) return;
    
    // Group products by location
    const locations = {};
    products.forEach(p => {
        const loc = (p.location || 'بێ ناو').toUpperCase().trim();
        if (!locations[loc]) locations[loc] = { name: loc, count: 0, products: [] };
        locations[loc].count++;
        locations[loc].products.push(p);
    });

    shelfGrid.innerHTML = '';
    
    if (Object.keys(locations).length === 0) {
        shelfGrid.innerHTML = '<div style="color:#6b7280; text-align:center; padding: 20px;">هیچ ڕەفەیەک نییە!</div>';
        return;
    }

    Object.values(locations).forEach(loc => {
        const div = document.createElement('div');
        div.className = `shelf-card ${activeSelectedShelf === loc.name ? 'active' : ''}`;
        div.style.cssText = `
            background: ${activeSelectedShelf === loc.name ? 'rgba(168,85,247,0.15)' : 'var(--card-bg)'};
            border: 2px solid ${activeSelectedShelf === loc.name ? '#a855f7' : 'var(--border-color)'};
            border-radius: 12px;
            padding: 15px;
            cursor: pointer;
            text-align: center;
            transition: all 0.2s;
            position: relative;
        `;
        div.innerHTML = `
            <div style="font-size:1.5rem; font-weight:700; color:var(--text-main); margin-bottom:5px;">${loc.name}</div>
            <div style="font-size:0.8rem; color:var(--text-dim);"><span style="color:#a855f7; font-weight:bold;">${loc.count}</span> بەرهەم</div>
            ${activeSelectedShelf === loc.name ? '<div style="position:absolute; top:-8px; right:-8px; background:#a855f7; color:#fff; border-radius:50%; width:24px; height:24px; display:flex; align-items:center; justify-content:center; font-size:0.8rem;">📍</div>' : ''}
        `;
        div.onclick = () => {
            activeSelectedShelf = activeSelectedShelf === loc.name ? null : loc.name;
            updateShelfLiveView();
        };
        shelfGrid.appendChild(div);
    });

    const panelContent = document.getElementById('shelfDetailContent');
    const panelTitle = document.getElementById('shelfDetailTitle');
    
    if (!panelContent || !panelTitle) return;

    if (!activeSelectedShelf) {
        panelTitle.innerText = 'وردەکاری ڕەفە';
        panelContent.innerHTML = '<div style="padding-top: 40px;">کلیک لەسەر ڕەفەیەک بکە بۆ بینینی بەرهەمەکانی ناوی...</div>';
    } else {
        const locData = locations[activeSelectedShelf];
        panelTitle.innerText = `📦 کاڵاکانی ناو ڕەفەی: ${activeSelectedShelf}`;
        if (!locData) return;
        
        let html = '<div style="display:flex; flex-direction:column; gap:8px; text-align:right; max-height:220px; overflow-y:auto; padding-right:5px;">';
        locData.products.forEach(p => {
            let status = p.qty > 0 ? '<span style="color:#10b981;font-size:0.75rem;">(مەوجودە)</span>' : '<span style="color:#ef4444;font-size:0.75rem;">(نەماوە)</span>';
            html += `
                <div style="background:var(--modal-input-bg); border:1px solid var(--border-color); padding:8px 12px; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:600; font-size:0.85rem; color:var(--text-main);">${p.name} <br> <span style="font-size:0.75rem; color:var(--text-dim); font-weight:normal;">${p.code}</span></div>
                    <div style="text-align:left;">
                        <div style="font-weight:700; color:#38bdf8; font-size:0.9rem;">${p.qty.toLocaleString()} <span style="font-size:0.7rem;">QTY</span></div>
                        ${status}
                    </div>
                </div>
            `;
        });
        html += '</div>';
        panelContent.innerHTML = html;
    }
}

function getFilteredProducts() {
    const filterName = document.getElementById('filterName') ? document.getElementById('filterName').value.toLowerCase().trim() : '';
    const filterBatch = document.getElementById('filterBatch') ? document.getElementById('filterBatch').value.toLowerCase().trim() : '';
    const filterCode = document.getElementById('filterCode') ? document.getElementById('filterCode').value.toLowerCase().trim() : '';
    const filterLocation = document.getElementById('filterLocation') ? document.getElementById('filterLocation').value.toLowerCase().trim() : '';
    const filterCompany = document.getElementById('filterCompany') ? document.getElementById('filterCompany').value.toLowerCase().trim() : '';
    
    const statusVal = invStatusFilter ? invStatusFilter.value : 'all';
    const sortVal = invSortFilter ? invSortFilter.value : 'name_asc';

    let filtered = [...products];

    // 1. Search Filters
    if (filterName) filtered = filtered.filter(p => (p.name || '').toLowerCase().includes(filterName));
    if (filterBatch) filtered = filtered.filter(p => (p.batchNo || '').toLowerCase().includes(filterBatch));
    if (filterCode) filtered = filtered.filter(p => (p.code || '').toLowerCase().includes(filterCode));
    if (filterLocation) filtered = filtered.filter(p => (p.location || '').toLowerCase().includes(filterLocation));
    if (filterCompany && filterCompany !== 'all') filtered = filtered.filter(p => (p.company || '').toLowerCase().includes(filterCompany));

    // 2. Status Filter
    const now = new Date();
    if (statusVal === 'in_stock') {
        filtered = filtered.filter(p => p.qty > 0 && p.stickerQty >= 500);
    } else if (statusVal === 'low_stock') {
        filtered = filtered.filter(p => p.qty <= 5 || p.stickerQty < 500);
    } else if (statusVal === 'out_of_stock') {
        filtered = filtered.filter(p => p.qty === 0 || p.stickerQty === 0);
    } else if (statusVal === 'exp_soon' || statusVal === 'soon') {
        filtered = filtered.filter(p => {
            if (!p.expDate) return false;
            const exp = new Date(p.expDate);
            const diffMonths = (exp - now) / (1000 * 60 * 60 * 24 * 30);
            return diffMonths < 6 && diffMonths >= 0;
        });
    } else if (statusVal === 'expired') {
        filtered = filtered.filter(p => new Date(p.expDate) < now);
    } else if (statusVal === 'ok') {
        filtered = filtered.filter(p => {
            const exp = new Date(p.expDate);
            const diffMonths = (exp - now) / (1000 * 60 * 60 * 24 * 30);
            return diffMonths >= 6;
        });
    }

    // 4. Sorting Logic
    filtered.sort((a, b) => {
        if (sortVal === 'name' || sortVal === 'name_asc') return (a.name || '').localeCompare(b.name || '');
        if (sortVal === 'name_desc') return (b.name || '').localeCompare(a.name || '');
        if (sortVal === 'qty_desc') return (b.qty || 0) - (a.qty || 0);
        if (sortVal === 'qty_asc') return (a.qty || 0) - (b.qty || 0);
        if (sortVal === 'sticker_desc') return (b.stickerQty || 0) - (a.stickerQty || 0);
        if (sortVal === 'exp_asc') {
            if (!a.expDate) return 1;
            if (!b.expDate) return -1;
            return new Date(a.expDate) - new Date(b.expDate);
        }
        if (sortVal === 'exp_desc') {
            if (!a.expDate) return 1;
            if (!b.expDate) return -1;
            return new Date(b.expDate) - new Date(a.expDate);
        }
        if (sortVal === 'price_desc') return (b.price || 0) - (a.price || 0);
        if (sortVal === 'received_desc') return new Date(b.receivedDate) - new Date(a.receivedDate);
        return 0;
    });

    return filtered;
}

function updateCompanyDropdown() {
    if (!invCompanyFilter) return;
    
    // If it's a text input, we don't populate dropdown options
    if (invCompanyFilter.tagName && invCompanyFilter.tagName.toLowerCase() === 'input') {
        return;
    }

    const currentSelection = invCompanyFilter.value;

    // Extract unique companies
    const companies = [...new Set(products.map(p => p.company).filter(c => c))];

    invCompanyFilter.innerHTML = '<option value="all">هەموو کۆمپانیاکان</option>';
    companies.forEach(c => {
        const option = document.createElement('option');
        option.value = c;
        option.textContent = c;
        invCompanyFilter.appendChild(option);
    });

    // Keep selection if it still exists
    if (companies.includes(currentSelection)) {
        invCompanyFilter.value = currentSelection;
    } else {
        invCompanyFilter.value = 'all';
    }
}

function renderProfessionalInventory() {
    // 1. Calculate Inventory Statistics
    const totalCount = products.length;
    const totalStickers = products.reduce((acc, p) => acc + (p.stickerQty || 0), 0);
    const totalValue = products.reduce((acc, p) => acc + (p.price * p.qty), 0);

    // Low Stock Alert Count (Sticker Stock < 500 or Qty <= 5)
    const lowStockCount = products.filter(p => p.qty <= 5 || p.stickerQty < 500).length;

    // Update UI Stats Cards
    const productsEl = document.getElementById('invTotalProducts');
    const stickersEl = document.getElementById('invTotalStickers');
    const valueEl = document.getElementById('invTotalValue');
    const lowEl = document.getElementById('invLowStockAlerts');

    if (productsEl) productsEl.innerText = totalCount.toLocaleString();
    if (stickersEl) stickersEl.innerText = totalStickers.toLocaleString();
    if (valueEl) valueEl.innerText = totalValue.toLocaleString() + ' IQD';
    if (lowEl) lowEl.innerText = lowStockCount.toLocaleString();

    // 2. Load dynamic company list
    updateCompanyDropdown();

    // 3. Filter and Sort products
    const filtered = getFilteredProducts();

    // Update the search count UI for sectionDashboard
    const resultNumEl = document.getElementById('filterResultNum');
    if (resultNumEl) {
        resultNumEl.innerText = filtered.length;
    }
    const noResultsRow = document.getElementById('noFilterResults');
    if (noResultsRow) {
        noResultsRow.style.display = filtered.length === 0 ? 'block' : 'none';
    }

    // 4. Update Shelf Live View
    updateShelfLiveView();

    // 5. Render Active View
    if (currentInventoryView === 'table') {
        if (!invTableBody) return;
        invTableBody.innerHTML = '';

        if (filtered.length === 0) {
            // Keep it empty, the noFilterResults element will show.
            return;
        }

        filtered.forEach(p => {
            const exp = new Date(p.expDate);
            const now = new Date();
            const diffMonths = (exp - now) / (1000 * 60 * 60 * 24 * 30);

            let expClass = 'in-stock';
            let expStatus = 'سەلامەتە ✅';
            if (diffMonths < 0) {
                expClass = 'low-stock';
                expStatus = 'بەسەرچووە ❌';
            } else if (diffMonths < 6) {
                expClass = 'exp-soon';
                expStatus = 'بەم زووانە بەسەردەچێت ⚠️';
            }

            // Qty level check
            let qtyClass = 'in-stock';
            if (p.qty === 0) qtyClass = 'low-stock';
            else if (p.qty <= 5) qtyClass = 'exp-soon';

            // Sticker Qty warning level check
            let stickerClass = 'in-stock';
            if (p.stickerQty === 0) stickerClass = 'low-stock';
            else if (p.stickerQty < 500) stickerClass = 'exp-soon';

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><span class="batch-badge">${p.batchNo || '-'}</span></td>
                <td><code class="code-cell">${p.code || '-'}</code></td>
                <td class="product-name-cell" style="font-weight:600; color: #1f2937;">${p.name}</td>
                <td><span class="stock-badge ${qtyClass}" style="font-weight:700;">${p.qty.toLocaleString()} QTY</span></td>
                <td>
                    <div style="display:flex; flex-direction:column; gap:4px;">
                        <span class="stock-badge ${stickerClass}" style="font-weight:700; width:fit-content;">${p.stickerQty.toLocaleString()} ستیکەر</span>
                        <div style="width: 80px; height: 4px; background: rgba(0,0,0,0.05); border-radius: 2px; overflow: hidden; margin-top:2px;">
                            <div style="width: ${Math.min(100, Math.max(0, (p.stickerQty / 15000) * 100))}%; height: 100%; background: ${p.stickerQty < 500 ? '#ef4444' : '#10b981'};"></div>
                        </div>
                    </div>
                </td>
                <td style="color:#94a3b8;">${p.price.toLocaleString()} د.ع</td>
                <td style="font-weight:700; color:#10b981;">${(p.qty * p.price).toLocaleString()} د.ع</td>
                <td><span class="location-badge">${p.location || '-'}</span></td>
                <td style="color:#94a3b8; font-size:0.85rem;">${p.company || '-'}</td>
                <td>
                    <span class="stock-badge ${expClass}" style="font-size:0.75rem; display:inline-block; line-height:1.2; text-align:center;">
                        ${formatDate(p.expDate)}<br/>
                        <small style="opacity:0.85; font-size:0.65rem;">${expStatus}</small>
                    </span>
                </td>
                <td>
                    <div class="action-btns">
                        <button class="action-btn sell" title="ڕێکخستنی خێرا (Adjust)" onclick="quickAdjustStock(${p.id})" style="background: rgba(56, 189, 248, 0.1); color: #38bdf8; border-color: rgba(56, 189, 248, 0.2);">🔄</button>
                        <button class="action-btn edit" title="دەستکاری" onclick="editProduct(${p.id})">✏️</button>
                        <button class="action-btn delete" title="سڕینەوە" onclick="deleteProduct(${p.id})">🗑️</button>
                    </div>
                </td>
            `;
            invTableBody.appendChild(tr);
        });
    } else {
        if (!invGridView) return;
        invGridView.innerHTML = '';

        if (filtered.length === 0) {
            invGridView.innerHTML = '<div class="no-search-msg" style="grid-column: 1/-1; text-align: center; padding: 50px; color: #ef4444;"><p style="font-size: 1.2rem;">❌ هیچ بەرهەمێک نەدۆزرایەوە بەپێی ئەم فلتەرانە!</p></div>';
            return;
        }

        filtered.forEach(p => {
            const exp = new Date(p.expDate);
            const now = new Date();
            const diffMonths = (exp - now) / (1000 * 60 * 60 * 24 * 30);

            let expClass = 'in-stock';
            let expStatus = 'تەندروستە ✅';
            if (diffMonths < 0) {
                expClass = 'low-stock';
                expStatus = 'بەسەرچووە ❌';
            } else if (diffMonths < 6) {
                expClass = 'exp-soon';
                expStatus = 'بەم زووانە بەسەردەچێت ⚠️';
            }

            // Warning indicators
            const isLowStock = p.qty <= 5 || p.stickerQty < 500;

            const card = document.createElement('div');
            card.className = 'inventory-card';
            card.style.cssText = `
                position: relative; 
                background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); 
                border: 1px solid ${isLowStock ? 'rgba(239, 68, 68, 0.2)' : 'rgba(255, 255, 255, 0.05)'}; 
                box-shadow: 0 10px 20px rgba(0,0,0,0.25);
            `;
            card.innerHTML = `
                <div style="display:flex; justify-content:space-between; margin-bottom:12px; align-items:center;">
                    <span class="batch-badge" style="font-size:0.75rem;">باچ: ${p.batchNo || '-'}</span>
                    <span class="location-badge" style="margin:0;">${p.location || '-'}</span>
                </div>
                <h4 style="font-size:1.1rem; color: #1f2937; font-weight:700; margin-bottom:15px; line-height:1.4;">${p.name}</h4>
                
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:15px;">
                    <div style="background:rgba(255,255,255,0.02); padding:10px; border-radius:10px; text-align:center; border:1px solid rgba(255,255,255,0.03);">
                        <span style="font-size:0.75rem; color:#64748b; display:block; margin-bottom:4px;">بڕی کۆگا</span>
                        <strong style="font-size:1.1rem; color:#38bdf8;">${p.qty.toLocaleString()} QTY</strong>
                    </div>
                    <div style="background:rgba(255,255,255,0.02); padding:10px; border-radius:10px; text-align:center; border:1px solid rgba(255,255,255,0.03);">
                        <span style="font-size:0.75rem; color:#64748b; display:block; margin-bottom:4px;">بڕی ستیکەر</span>
                        <strong style="font-size:1.1rem; color:#fbbf24;">${p.stickerQty.toLocaleString()}</strong>
                    </div>
                </div>

                <div style="margin-bottom:15px;">
                    <div style="display:flex; justify-content:space-between; font-size:0.75rem; margin-bottom:4px; color:#64748b;">
                        <span>ڕێژەی کۆی بەها</span>
                        <span style="color:#10b981; font-weight:600;">${(p.qty * p.price).toLocaleString()} د.ع</span>
                    </div>
                    <div style="width:100%; height:6px; background:rgba(0,0,0,0.05); border-radius:4px; overflow:hidden;">
                        <div style="width: ${Math.min(100, Math.max(0, (p.stickerQty / 15000) * 100))}%; height:100%; background:linear-gradient(90deg, #fbbf24, #f59e0b); border-radius:4px;"></div>
                    </div>
                </div>

                <div style="display:flex; justify-content:space-between; align-items:center; border-top:1px solid rgba(0,0,0,0.05); padding-top:12px; margin-top:10px; font-size:0.8rem;">
                    <div>
                        <span style="font-size:0.7rem; color:#64748b; display:block;">بەواری بەسەرچوون</span>
                        <span class="stock-badge ${expClass}" style="padding:2px 8px; font-size:0.72rem; font-weight:700;">${formatDate(p.expDate)} (${expStatus})</span>
                    </div>
                    <div class="action-btns" style="margin-top:0;">
                        <button class="action-btn sell" title="ڕێکخستنی خێرا (Adjust)" onclick="quickAdjustStock(${p.id})" style="background: rgba(56, 189, 248, 0.1); color: #38bdf8; border-color: rgba(56, 189, 248, 0.2);">🔄</button>
                        <button class="action-btn edit" title="دەستکاری" onclick="editProduct(${p.id})">✏️</button>
                        <button class="action-btn delete" title="سڕینەوە" onclick="deleteProduct(${p.id})">🗑️</button>
                    </div>
                </div>
            `;
            invGridView.appendChild(card);
        });
    }
    if (window.renderWarehouseShelfMap) window.renderWarehouseShelfMap();
}

function quickAdjustStock(id) {
    const p = products.find(prod => prod.id === id);
    if (!p) return;

    // Quick adjustment prompt for stickers
    const newStickers = prompt(`🔄 دەستکاری بڕی ستیکەر بۆ بەرهەمی:\n"${p.name}"\n\nبڕی ئێستا: ${p.stickerQty.toLocaleString()}`, p.stickerQty);
    if (newStickers === null) return;

    const qtyVal = parseInt(newStickers.replace(/,/g, ''));
    if (isNaN(qtyVal) || qtyVal < 0) {
        showToast('⚠️ تکایە بڕێکی دروست داخڵ بکە!');
        return;
    }

    // Quick adjustment prompt for Qty
    const newQty = prompt(`🔄 دەستکاری بڕی مەوجودی کۆگا (QTY) بۆ بەرهەمی:\n"${p.name}"\n\nبڕی ئێستا: ${p.qty.toLocaleString()}`, p.qty);
    if (newQty === null) return;

    const qtyProd = parseInt(newQty.replace(/,/g, ''));
    if (isNaN(qtyProd) || qtyProd < 0) {
        showToast('⚠️ تکایە بڕێکی دروست داخڵ بکە!');
        return;
    }

    // Apply changes
    p.stickerQty = qtyVal;
    p.qty = qtyProd;

    saveData('sticker_products', products);

    // Re-render
    renderTable();
    renderProfessionalInventory();
    updateSalesDropdown();

    showToast('🔄 بڕەکان بە سەرکەوتوویی نوێکرانەوە!');
}

// ==================== REPORTS LOGIC ====================
let workerProductivityChart = null;
let inventoryShareChart = null;
let stickerSalesTrendsChart = null;
let dashboardDailySalesChart = null;
let dashboardSalesWasteChart = null;

function renderAnalyticsCharts() {
    // 1. Worker Productivity Chart (Bar Chart)
    const workerCtx = document.getElementById('workerProductivityChart');
    if (workerCtx) {
        const now = new Date();
        const currentMonth = now.getMonth();
        const currentYear = now.getFullYear();
        const workerStats = {};

        sales.forEach(sale => {
            const saleDate = new Date(sale.date);
            if (saleDate.getMonth() === currentMonth && saleDate.getFullYear() === currentYear) {
                (sale.assignedWorkers || []).forEach(w => {
                    if (!workerStats[w.id]) {
                        workerStats[w.id] = { name: w.name, totalStickers: 0 };
                    }
                    workerStats[w.id].totalStickers += w.qty;
                });
            }
        });

        const workerLabels = [];
        const workerData = [];
        Object.values(workerStats)
            .sort((a, b) => b.totalStickers - a.totalStickers)
            .forEach(w => {
                workerLabels.push(w.name);
                workerData.push(w.totalStickers);
            });

        if (workerProductivityChart) workerProductivityChart.destroy();

        workerProductivityChart = new Chart(workerCtx, {
            type: 'bar',
            data: {
                labels: workerLabels.length > 0 ? workerLabels : ['هیچ کڕێکارێک'],
                datasets: [{
                    label: 'کۆی ستیکەری لێدراو (دانە)',
                    data: workerData.length > 0 ? workerData : [0],
                    backgroundColor: 'rgba(168, 85, 247, 0.65)',
                    borderColor: '#a855f7',
                    borderWidth: 2,
                    borderRadius: 8,
                    hoverBackgroundColor: 'rgba(168, 85, 247, 0.85)',
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: (context) => ` ${context.parsed.y.toLocaleString()} ستیکەر`
                        }
                    }
                },
                scales: {
                    x: {
                        grid: { color: 'rgba(0,0,0,0.05)' },
                        ticks: { color: '#94a3b8', font: { family: 'Outfit' } }
                    },
                    y: {
                        grid: { color: 'rgba(0,0,0,0.05)' },
                        ticks: { color: '#94a3b8', font: { family: 'Outfit' } }
                    }
                }
            }
        });
    }

    // 2. Inventory Distribution Chart (Doughnut Chart)
    const invCtx = document.getElementById('inventoryShareChart');
    if (invCtx) {
        const sortedProducts = [...products].sort((a, b) => (b.stickerQty || 0) - (a.stickerQty || 0));

        const topN = 5;
        const labels = [];
        const data = [];
        let othersSum = 0;

        sortedProducts.forEach((p, idx) => {
            if (idx < topN) {
                labels.push(p.name);
                data.push(p.stickerQty || 0);
            } else {
                othersSum += p.stickerQty || 0;
            }
        });

        if (sortedProducts.length > topN && othersSum > 0) {
            labels.push('کاڵاکانی تر');
            data.push(othersSum);
        }

        if (inventoryShareChart) inventoryShareChart.destroy();

        inventoryShareChart = new Chart(invCtx, {
            type: 'doughnut',
            data: {
                labels: labels.length > 0 ? labels : ['هیچ کاڵایەک'],
                datasets: [{
                    data: data.length > 0 ? data : [0],
                    backgroundColor: [
                        'rgba(56, 189, 248, 0.7)',
                        'rgba(16, 185, 129, 0.7)',
                        'rgba(251, 191, 36, 0.7)',
                        'rgba(239, 68, 68, 0.7)',
                        'rgba(139, 92, 246, 0.7)',
                        'rgba(100, 116, 139, 0.7)'
                    ],
                    borderColor: '#1e293b',
                    borderWidth: 2,
                    hoverOffset: 10
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'right',
                        labels: {
                            color: '#94a3b8',
                            font: { family: 'Outfit', size: 11 },
                            padding: 15
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: (context) => ` ${context.parsed.toLocaleString()} ستیکەر`
                        }
                    }
                }
            }
        });
    }

    // 3. Monthly Sticker Sales Trends (Line Chart)
    const salesCtx = document.getElementById('stickerSalesTrendsChart');
    if (salesCtx) {
        const now = new Date();
        const currentMonth = now.getMonth();
        const currentYear = now.getFullYear();

        const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
        const dailyTotals = Array(daysInMonth).fill(0);

        sales.forEach(sale => {
            const saleDate = new Date(sale.date);
            if (saleDate.getMonth() === currentMonth && saleDate.getFullYear() === currentYear) {
                const day = saleDate.getDate();
                const stickerSum = (sale.assignedWorkers || []).reduce((sum, w) => sum + w.qty, 0);
                dailyTotals[day - 1] += stickerSum;
            }
        });

        const dayLabels = Array.from({ length: daysInMonth }, (_, i) => `${i + 1}`);

        if (stickerSalesTrendsChart) stickerSalesTrendsChart.destroy();

        stickerSalesTrendsChart = new Chart(salesCtx, {
            type: 'line',
            data: {
                labels: dayLabels,
                datasets: [{
                    label: 'ستیکەری فرۆشراو بەپێی ڕۆژ',
                    data: dailyTotals,
                    fill: true,
                    backgroundColor: 'rgba(16, 185, 129, 0.12)',
                    borderColor: '#10b981',
                    borderWidth: 3,
                    tension: 0.4,
                    pointBackgroundColor: '#10b981',
                    pointBorderColor: '#0f172a',
                    pointBorderWidth: 2,
                    pointRadius: 4,
                    pointHoverRadius: 6,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: (context) => ` ${context.parsed.y.toLocaleString()} ستیکەر`
                        }
                    }
                },
                scales: {
                    x: {
                        grid: { color: 'rgba(255,255,255,0.02)' },
                        ticks: { color: '#64748b', font: { family: 'Outfit', size: 10 } }
                    },
                    y: {
                        grid: { color: 'rgba(0,0,0,0.05)' },
                        ticks: { color: '#94a3b8', font: { family: 'Outfit' } }
                    }
                }
            }
        });
    }

    // 4. Daily Sales Trend (Line Chart - Compact)
    const dailySalesCtx = document.getElementById('dashboardDailySalesChart');
    if (dailySalesCtx) {
        const now = new Date();
        const currentMonth = now.getMonth();
        const currentYear = now.getFullYear();

        const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
        const dailyTotals = Array(daysInMonth).fill(0);

        sales.forEach(sale => {
            const saleDate = new Date(sale.date);
            if (saleDate.getMonth() === currentMonth && saleDate.getFullYear() === currentYear) {
                const day = saleDate.getDate();
                const stickerSum = (sale.assignedWorkers || []).reduce((sum, w) => sum + w.qty, 0);
                dailyTotals[day - 1] += stickerSum;
            }
        });

        const dayLabels = Array.from({ length: daysInMonth }, (_, i) => `${i + 1}`);

        if (dashboardDailySalesChart) dashboardDailySalesChart.destroy();

        dashboardDailySalesChart = new Chart(dailySalesCtx, {
            type: 'line',
            data: {
                labels: dayLabels,
                datasets: [{
                    label: 'ستیکەری فرۆشراو بەپێی ڕۆژ',
                    data: dailyTotals,
                    fill: true,
                    backgroundColor: 'rgba(16, 185, 129, 0.12)',
                    borderColor: '#10b981',
                    borderWidth: 2,
                    tension: 0.4,
                    pointBackgroundColor: '#10b981',
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 1.5,
                    pointRadius: 3,
                    pointHoverRadius: 5,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: (context) => ` ${context.parsed.y.toLocaleString()} ستیکەر`
                        }
                    }
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { color: '#94a3b8', font: { family: 'Outfit', size: 9 } }
                    },
                    y: {
                        grid: { color: 'rgba(0,0,0,0.05)' },
                        ticks: { color: '#94a3b8', font: { family: 'Outfit', size: 9 } }
                    }
                }
            }
        });
    }

    // 5. Sales vs Waste (Doughnut Chart - Compact)
    const salesWasteCtx = document.getElementById('dashboardSalesWasteChart');
    if (salesWasteCtx) {
        const now = new Date();
        const currentMonth = now.getMonth();
        const currentYear = now.getFullYear();

        let totalSalesStickers = 0;
        let totalWasteStickers = 0;

        sales.forEach(sale => {
            const saleDate = new Date(sale.date);
            if (saleDate.getMonth() === currentMonth && saleDate.getFullYear() === currentYear) {
                const stickerSum = (sale.assignedWorkers || []).reduce((sum, w) => sum + w.qty, 0);
                totalSalesStickers += stickerSum;
            }
        });

        talaf.forEach(t => {
            const wasteDate = new Date(t.date);
            if (wasteDate.getMonth() === currentMonth && wasteDate.getFullYear() === currentYear) {
                totalWasteStickers += t.stickerQty || 0;
            }
        });

        if (dashboardSalesWasteChart) dashboardSalesWasteChart.destroy();

        dashboardSalesWasteChart = new Chart(salesWasteCtx, {
            type: 'doughnut',
            data: {
                labels: ['فرۆشتن', 'تکجوو (وێرانبوو)'],
                datasets: [{
                    data: [totalSalesStickers, totalWasteStickers],
                    backgroundColor: [
                        'rgba(16, 185, 129, 0.75)', // Green for sales
                        'rgba(239, 68, 68, 0.75)'  // Red for waste/talaf
                    ],
                    borderColor: '#ffffff',
                    borderWidth: 1.5,
                    hoverOffset: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'right',
                        labels: {
                            color: '#64748b',
                            font: { family: 'Outfit', size: 10 },
                            padding: 8
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: (context) => ` ${context.label}: ${context.parsed.toLocaleString()} ستیکەر`
                        }
                    }
                }
            }
        });
    }
}

// ==================== REPORTS LOGIC ====================
function renderWorkerReport() {
    if (!workerReportTableBody) return;
    workerReportTableBody.innerHTML = '';

    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    // Stats map: { workerId: { name: string, totalStickers: number, jobsCount: number, lastDate: string } }
    const workerStats = {};

    sales.forEach(sale => {
        const saleDate = new Date(sale.date);
        if (saleDate.getMonth() === currentMonth && saleDate.getFullYear() === currentYear) {
            (sale.assignedWorkers || []).forEach(w => {
                if (!workerStats[w.id]) {
                    workerStats[w.id] = {
                        name: w.name,
                        totalStickers: 0,
                        jobsCount: 0,
                        lastDate: sale.date
                    };
                }
                workerStats[w.id].totalStickers += w.qty;
                workerStats[w.id].jobsCount += 1;
                if (new Date(sale.date) > new Date(workerStats[w.id].lastDate)) {
                    workerStats[w.id].lastDate = sale.date;
                }
            });
        }
    });

    const statsArray = Object.values(workerStats).sort((a, b) => b.totalStickers - a.totalStickers);

    if (statsArray.length === 0) {
        workerReportTableBody.innerHTML = '<tr><td colspan="4" style="text-align: center; padding: 30px; color: #64748b;">هیچ چالاکییەک لەم مانگەدا تۆمار نەکراوە.</td></tr>';
        renderAnalyticsCharts();
        return;
    }

    statsArray.forEach(stat => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight: 600; color: #1f2937;">${stat.name}</td>
            <td style="color: #38bdf8; font-weight: 700;">${stat.totalStickers.toLocaleString()} ستیکەر</td>
            <td>${stat.jobsCount} ئیش</td>
            <td style="font-size: 0.8rem; color: #94a3b8;">${new Date(stat.lastDate).toLocaleString('ku-IQ')}</td>
        `;
        workerReportTableBody.appendChild(tr);
    });
    renderAnalyticsCharts();
}

// ==================== UTILS ====================
function formatDate(dateStr) {
    if (!dateStr) return '-';
    const d = new Date(dateStr);
    return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

function showToast(msg) {
    const toast = document.createElement('div');
    toast.className = 'toast show';
    toast.innerText = msg;
    document.body.appendChild(toast);
    setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 400); }, 3000);
}

// Initial Load
renderTable();
renderWorkerTable();
renderSalesTable();
updateSalesDropdown();
renderProfessionalInventory();

// ==================== PAYROLL / ATTENDANCE LOGIC ====================

// Declarations moved to the top of the file to prevent TDZ ReferenceErrors on startup

function renderQuickClockList() {
    const listContainer = document.getElementById('quickClockList');
    if (!listContainer) return;

    listContainer.innerHTML = `
        <style>
        @keyframes clockPulse {
            0% { opacity: 0.4; }
            50% { opacity: 1; }
            100% { opacity: 0.4; }
        }
        .pulse-dot {
            width: 8px;
            height: 8px;
            background-color: #22c55e;
            border-radius: 50%;
            display: inline-block;
            animation: clockPulse 1.5s infinite ease-in-out;
        }
        .clock-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: rgba(255,255,255,0.02);
            border: 1px solid rgba(255,255,255,0.06);
            border-radius: 12px;
            padding: 12px 16px;
            transition: all 0.2s ease;
        }
        .clock-row:hover {
            background: rgba(0,0,0,0.04);
            border-color: rgba(0,0,0,0.1);
        }
        </style>
    `;

    if (workers.length === 0) {
        listContainer.innerHTML += '<div style="text-align: center; color: #64748b; padding: 20px 0;">هیچ کڕێکارێک تۆمار نەکراوە.</div>';
        return;
    }

    workers.forEach(w => {
        const isClockedIn = !!clockInLogs[w.id];
        const row = document.createElement('div');
        row.className = 'clock-row';

        if (isClockedIn) {
            const startTime = new Date(clockInLogs[w.id]);
            const timeStr = startTime.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
            row.innerHTML = `
                <div>
                    <h4 style="margin: 0; color: #1f2937; font-weight: 600;">${w.name}</h4>
                    <span style="font-size: 0.75rem; color: #64748b;">${w.job}</span>
                </div>
                <div style="display: flex; align-items: center; gap: 15px;">
                    <div style="display: flex; flex-direction: column; align-items: flex-end; gap: 4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; padding:3px 10px; background:rgba(34,197,94,0.15); border:1px solid rgba(34,197,94,0.3); border-radius:12px; color:#22c55e; font-size:0.75rem; font-weight:bold;">
                            <span class="pulse-dot"></span> لە دەوامدایە
                        </span>
                        <span style="font-size: 0.72rem; color: #94a3b8;">⏱️ چوونەژوور: ${timeStr}</span>
                    </div>
                    <button onclick="clockOutWorker(${w.id})" style="padding: 8px 14px; background: linear-gradient(135deg, #ef4444, #dc2626); color: white; border: none; border-radius: 8px; font-size: 0.8rem; font-weight: bold; cursor: pointer; box-shadow: 0 4px 10px rgba(239, 68, 68, 0.2); transition: all 0.2s;">
                        🔴 تەواوبوون
                    </button>
                </div>
            `;
        } else {
            row.innerHTML = `
                <div>
                    <h4 style="margin: 0; color: #1f2937; font-weight: 600;">${w.name}</h4>
                    <span style="font-size: 0.75rem; color: #64748b;">${w.job}</span>
                </div>
                <div style="display: flex; align-items: center; gap: 15px;">
                    <span style="padding: 3px 10px; background: rgba(0,0,0,0.04); border: 1px solid rgba(255,255,255,0.08); border-radius: 12px; color: #94a3b8; font-size: 0.75rem;">
                        خەوتوو 😴
                    </span>
                    <button onclick="clockInWorker(${w.id})" style="padding: 8px 14px; background: linear-gradient(135deg, #22c55e, #16a34a); color: white; border: none; border-radius: 8px; font-size: 0.8rem; font-weight: bold; cursor: pointer; box-shadow: 0 4px 10px rgba(34, 197, 94, 0.2); transition: all 0.2s;">
                        🟢 دەستپێکردن
                    </button>
                </div>
            `;
        }
        listContainer.appendChild(row);
    });
}

function clockInWorker(id) {
    const w = workers.find(work => work.id == id);
    if (!w) return;

    clockInLogs[id] = Date.now();
    saveData('sticker_clock_in_logs', clockInLogs);
    renderQuickClockList();
    showToast(`🟢 دەستپێکردنی دەوامی کڕێکار ${w.name} تۆمار کرا!`);
}

function clockOutWorker(id) {
    const w = workers.find(work => work.id == id);
    if (!w) return;

    const startTime = clockInLogs[id];
    if (!startTime) return;

    const elapsedMs = Date.now() - startTime;
    let elapsedHours = 0;

    // Developer helper for quick demo clicks
    if (elapsedMs < 5 * 60 * 1000) {
        elapsedHours = 8.5; // Simulate a full shift + 0.5 hours overtime
    } else {
        elapsedHours = Math.round((elapsedMs / (1000 * 60 * 60)) * 2) / 2; // round to nearest 0.5 hour
    }

    // Pre-fill the Add Attendance form with calculated metrics
    const selectEl = document.getElementById('attendanceWorkerSelect');
    if (selectEl) {
        selectEl.value = id;
    }

    const dateEl = document.getElementById('attendanceDate');
    if (dateEl) {
        dateEl.value = new Date().toISOString().split('T')[0];
    }

    // Calculate Context-Aware Daily Base wage
    const dailyWage = w.salary < 80000 ? w.salary : Math.round(w.salary / 26);
    const dailyRateEl = document.getElementById('attendanceDailyRate');
    if (dailyRateEl) {
        dailyRateEl.value = dailyWage;
    }

    // Calculate Overtime (Standard shift is 8 hours)
    const overtimeHours = Math.max(0, elapsedHours - 8);
    const hoursEl = document.getElementById('attendanceHours');
    if (hoursEl) {
        hoursEl.value = overtimeHours;
    }

    const hourlyRateEl = document.getElementById('attendanceHourlyRate');
    if (hourlyRateEl) {
        hourlyRateEl.value = 3000; // standard hourly rate
    }

    const notesEl = document.getElementById('attendanceNotes');
    if (notesEl) {
        notesEl.value = `کاتی گشتی دەوام: ${elapsedHours} سەعات (تۆماری ئۆتۆماتیکی دەفتەری ئەلیکترۆنی)`;
    }

    // Update live pay preview boxes
    updatePayPreview();

    // Scroll smoothly to form
    const formEl = document.getElementById('attendanceForm');
    if (formEl) {
        formEl.scrollIntoView({ behavior: 'smooth' });
    }

    // Clean up active log session
    delete clockInLogs[id];
    saveData('sticker_clock_in_logs', clockInLogs);
    renderQuickClockList();

    showToast(`⏱️ دەوامی ${w.name} تەواو بوو! زانیارییەکان حیساب کران و ڕەوانەی فۆرمەکە کران.`);
}

window.clockInWorker = clockInWorker;
window.clockOutWorker = clockOutWorker;
window.renderQuickClockList = renderQuickClockList;

// --- Populate worker dropdowns for payroll ---
function populatePayrollWorkerDropdowns() {
    const sel = document.getElementById('attendanceWorkerSelect');
    const filterSel = document.getElementById('payrollRecordWorkerFilter');
    if (!sel || !filterSel) return;

    sel.innerHTML = '<option value="">هەڵبژێرە...</option>';
    filterSel.innerHTML = '<option value="">هەموو کڕێکارەکان</option>';

    workers.forEach(w => {
        const opt1 = document.createElement('option');
        opt1.value = w.id;
        opt1.textContent = w.name;
        sel.appendChild(opt1);

        const opt2 = document.createElement('option');
        opt2.value = w.id;
        opt2.textContent = w.name;
        filterSel.appendChild(opt2);
    });
}

// --- Set today's date as default ---
function initPayrollDate() {
    const dateInput = document.getElementById('attendanceDate');
    if (dateInput && !dateInput.value) {
        dateInput.value = new Date().toISOString().split('T')[0];
    }
}

// --- Live pay preview ---
function updatePayPreview() {
    const hours = parseFloat(document.getElementById('attendanceHours').value) || 0;
    const hourlyRate = parseFloat(document.getElementById('attendanceHourlyRate').value) || 0;
    const dailyRate = parseFloat(document.getElementById('attendanceDailyRate').value) || 0;
    const hourlyPay = hours * hourlyRate;
    const total = hourlyPay + dailyRate;
    const box = document.getElementById('payPreviewBox');
    if (hours > 0 || hourlyRate > 0 || dailyRate > 0) {
        box.style.display = 'block';
        document.getElementById('payPreviewHourly').textContent = hourlyPay.toLocaleString() + ' د.ع';
        document.getElementById('payPreviewDaily').textContent = dailyRate.toLocaleString() + ' د.ع';
        document.getElementById('payPreviewCalc').textContent = total.toLocaleString() + ' د.ع';
    } else {
        box.style.display = 'none';
    }
}
document.getElementById('attendanceHours').addEventListener('input', updatePayPreview);
document.getElementById('attendanceHourlyRate').addEventListener('input', updatePayPreview);
document.getElementById('attendanceDailyRate').addEventListener('input', updatePayPreview);

// --- Submit attendance form ---
document.getElementById('attendanceForm').addEventListener('submit', e => {
    e.preventDefault();
    const workerId = document.getElementById('attendanceWorkerSelect').value;
    const worker = workers.find(w => w.id == workerId);
    if (!worker) { showToast('⚠️ تکایە کڕێکارێک هەڵبژێرە!'); return; }

    const hours = parseFloat(document.getElementById('attendanceHours').value) || 0;
    const hourlyRate = parseInt(document.getElementById('attendanceHourlyRate').value) || 3000;
    const dailyRate = parseInt(document.getElementById('attendanceDailyRate').value) || 25000;
    const notes = document.getElementById('attendanceNotes').value.trim();
    const date = document.getElementById('attendanceDate').value;
    const hourlyPay = hours * hourlyRate;
    const totalPay = hourlyPay + dailyRate;

    const record = {
        id: Date.now(),
        workerId: worker.id,
        workerName: worker.name,
        date,
        hours,
        hourlyRate,
        dailyRate,
        hourlyPay,
        totalPay,
        notes
    };

    attendance.unshift(record);
    saveData('sticker_attendance', attendance);

    renderAttendanceTable();
    renderPayrollSummary();
    updatePayrollStats();

    document.getElementById('attendanceForm').reset();
    document.getElementById('attendanceHours').value = 0;
    document.getElementById('attendanceHourlyRate').value = 3000;
    document.getElementById('attendanceDailyRate').value = 25000;
    initPayrollDate();
    document.getElementById('payPreviewBox').style.display = 'none';
    showToast('✅ تۆمارەکە بە سەرکەوتوویی زیادکرا!');
});

// --- Render all-records table (with optional filters) ---
function renderAttendanceTable() {
    const tbody = document.getElementById('attendanceTableBody');
    if (!tbody) return;

    const wFilter = document.getElementById('payrollRecordWorkerFilter').value;
    const dFilter = document.getElementById('payrollRecordDateFilter').value;

    let filtered = attendance;
    if (wFilter) filtered = filtered.filter(r => r.workerId == wFilter);
    if (dFilter) filtered = filtered.filter(r => r.date === dFilter);

    tbody.innerHTML = '';
    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:30px;color:#64748b;">هیچ تۆمارێک نییە.</td></tr>';
        return;
    }

    filtered.forEach(r => {
        const hourlyRate = r.hourlyRate || (r.payPerLoad ? 0 : 3000);
        const dailyRate = r.dailyRate || (r.payPerLoad ? r.payPerLoad : 25000);
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight:600;">${r.workerName}</td>
            <td>${formatDate(r.date)}</td>
            <td><span style="color:#fbbf24;font-weight:700;">${r.hours} سەعات</span></td>
            <td style="color:#94a3b8;">${hourlyRate.toLocaleString()} د.ع</td>
            <td style="color:#94a3b8;">${dailyRate.toLocaleString()} د.ع</td>
            <td><strong style="color:#22c55e;">${r.totalPay.toLocaleString()} د.ع</strong></td>
            <td style="color:#94a3b8;font-size:.85rem;">${r.notes || '-'}</td>
            <td>
                <div class="action-btns">
                    <button class="action-btn delete" title="سڕینەوە" onclick="deleteAttendance(${r.id})">🗑️</button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// --- Delete a record ---
function deleteAttendance(id) {
    if (!confirm('دڵنیایت لە سڕینەوەی ئەم تۆمارە؟')) return;
    attendance = attendance.filter(r => r.id !== id);
    saveData('sticker_attendance', attendance);
    renderAttendanceTable();
    renderPayrollSummary();
    updatePayrollStats();
    showToast('تۆمارەکە سڕایەوە! 🗑️');
}

// --- Clear filters ---
function clearPayrollFilters() {
    document.getElementById('payrollRecordWorkerFilter').value = '';
    document.getElementById('payrollRecordDateFilter').value = '';
    renderAttendanceTable();
}

// --- Weekly summary table ---
function renderPayrollSummary() {
    const tbody = document.getElementById('payrollSummaryBody');
    const grandEl = document.getElementById('payrollGrandTotal');
    if (!tbody) return;

    const weekOffset = parseInt(document.getElementById('payrollWeekFilter').value) || 0;
    const now = new Date();

    let from, to;
    if (weekOffset === 4) {
        // Current month
        from = new Date(now.getFullYear(), now.getMonth(), 1);
        to = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    } else {
        // Week-based: Sunday = start of current week
        const dayOfWeek = now.getDay(); // 0=Sun
        const startOfThisWeek = new Date(now);
        startOfThisWeek.setDate(now.getDate() - dayOfWeek);
        startOfThisWeek.setHours(0, 0, 0, 0);

        from = new Date(startOfThisWeek);
        from.setDate(from.getDate() - weekOffset * 7);
        to = new Date(from);
        to.setDate(to.getDate() + 6);
        to.setHours(23, 59, 59, 999);
    }

    const filtered = attendance.filter(r => {
        const d = new Date(r.date);
        return d >= from && d <= to;
    });

    // Aggregate by worker
    const stats = {};
    filtered.forEach(r => {
        if (!stats[r.workerId]) {
            stats[r.workerId] = { name: r.workerName, days: new Set(), hours: 0, hourlyPay: 0, dailyPay: 0, totalPay: 0 };
        }
        stats[r.workerId].days.add(r.date);
        stats[r.workerId].hours += r.hours;
        stats[r.workerId].hourlyPay += (r.hourlyPay !== undefined ? r.hourlyPay : r.hours * (r.hourlyRate || 3000));
        stats[r.workerId].dailyPay += (r.dailyRate !== undefined ? r.dailyRate : (r.payPerLoad || 25000));
        stats[r.workerId].totalPay += r.totalPay;
    });

    tbody.innerHTML = '';
    const arr = Object.values(stats).sort((a, b) => b.totalPay - a.totalPay);

    if (arr.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:25px;color:#64748b;">هیچ تۆمارێک لەم ماوەیەدا نییە.</td></tr>';
        grandEl.textContent = '0 د.ع';
        return;
    }

    let grand = 0;
    arr.forEach(s => {
        grand += s.totalPay;
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight:600;color: #1f2937;">${s.name}</td>
            <td><span class="location-badge">${s.days.size} ڕۆژ</span></td>
            <td style="color:#fbbf24;font-weight:600;">${s.hours} سەعات</td>
            <td style="color:#38bdf8;">${s.hourlyPay.toLocaleString()} د.ع</td>
            <td style="color:#fbbf24;">${s.dailyPay.toLocaleString()} د.ع</td>
            <td><strong style="color:#22c55e;font-size:1rem;">${s.totalPay.toLocaleString()} د.ع</strong></td>
        `;
        tbody.appendChild(tr);
    });
    grandEl.textContent = grand.toLocaleString() + ' د.ع';
}

// --- Update today's stat cards ---
function updatePayrollStats() {
    const today = new Date().toISOString().split('T')[0];
    const todayRecs = attendance.filter(r => r.date === today);
    document.getElementById('payTodayCount').textContent = todayRecs.length;
    document.getElementById('payTodayHours').textContent = todayRecs.reduce((a, r) => a + r.hours, 0) + ' سەعات';
    const totalOTToday = todayRecs.reduce((a, r) => a + (r.hourlyPay !== undefined ? r.hourlyPay : r.hours * (r.hourlyRate || 3000)), 0);
    document.getElementById('payTodayLoads').textContent = totalOTToday.toLocaleString() + ' د.ع';
    const totalToday = todayRecs.reduce((a, r) => a + r.totalPay, 0);
    document.getElementById('payTodayTotal').textContent = totalToday.toLocaleString() + ' د.ع';
}

// --- Week filter change ---
document.getElementById('payrollWeekFilter').addEventListener('change', renderPayrollSummary);

// --- Record filters ---
document.getElementById('payrollRecordWorkerFilter').addEventListener('change', renderAttendanceTable);
document.getElementById('payrollRecordDateFilter').addEventListener('change', renderAttendanceTable);

// --- Payroll Excel Export ---
document.getElementById('exportPayrollBtn').addEventListener('click', () => {
    const weekOffset = parseInt(document.getElementById('payrollWeekFilter').value) || 0;
    const now = new Date();
    let from, to;
    if (weekOffset === 4) {
        from = new Date(now.getFullYear(), now.getMonth(), 1);
        to = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    } else {
        const dayOfWeek = now.getDay();
        const startOfThisWeek = new Date(now);
        startOfThisWeek.setDate(now.getDate() - dayOfWeek);
        startOfThisWeek.setHours(0, 0, 0, 0);
        from = new Date(startOfThisWeek);
        from.setDate(from.getDate() - weekOffset * 7);
        to = new Date(from);
        to.setDate(to.getDate() + 6);
        to.setHours(23, 59, 59, 999);
    }

    const filtered = attendance.filter(r => {
        const d = new Date(r.date); return d >= from && d <= to;
    });

    const wb = XLSX.utils.book_new();

    // Detail sheet
    const detailData = [
        ['کڕێکار', 'بەروار', 'سەعاتی کار', 'نرخی سەعات', 'مووچەی سەعات', 'مووچەی دیاری', 'کۆی مووچە', 'تێبینی']
    ];
    filtered.forEach(r => {
        const hRate = r.hourlyRate || 3000;
        const dRate = r.dailyRate || (r.payPerLoad || 25000);
        const hPay = r.hourlyPay !== undefined ? r.hourlyPay : r.hours * hRate;
        detailData.push([r.workerName, r.date, r.hours, hRate, hPay, dRate, r.totalPay, r.notes || '']);
    });
    const ws1 = XLSX.utils.aoa_to_sheet(detailData);
    ws1['!cols'] = [{ wch: 22 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 25 }];
    XLSX.utils.book_append_sheet(wb, ws1, 'تۆمارەکانی ڕوژانە');

    // Summary sheet
    const stats2 = {};
    filtered.forEach(r => {
        if (!stats2[r.workerId]) stats2[r.workerId] = { name: r.workerName, days: new Set(), hours: 0, hourlyPay: 0, dailyPay: 0, totalPay: 0 };
        stats2[r.workerId].days.add(r.date);
        stats2[r.workerId].hours += r.hours;
        stats2[r.workerId].hourlyPay += (r.hourlyPay !== undefined ? r.hourlyPay : r.hours * (r.hourlyRate || 3000));
        stats2[r.workerId].dailyPay += (r.dailyRate !== undefined ? r.dailyRate : (r.payPerLoad || 25000));
        stats2[r.workerId].totalPay += r.totalPay;
    });
    const summaryData = [['کڕێکار', 'ژمارەی ڕۆژەکان', 'کۆی سەعات', 'مووچەی سەعات', 'مووچەی دیاری', 'کۆی مووچە']];
    Object.values(stats2).forEach(s => {
        summaryData.push([s.name, s.days.size, s.hours, s.hourlyPay, s.dailyPay, s.totalPay]);
    });
    const grandT = Object.values(stats2).reduce((a, s) => a + s.totalPay, 0);
    summaryData.push(['', '', '', '', 'کۆی گشتی', grandT]);
    const ws2 = XLSX.utils.aoa_to_sheet(summaryData);
    ws2['!cols'] = [{ wch: 22 }, { wch: 16 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, ws2, 'کۆی مووچەی کڕێکارەکان');

    try {
        const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        const dateStr = new Date().toISOString().split('T')[0];
        const fileName = `Payroll_${dateStr}.xlsx`;
        const savePath = electronDialog
            ? electronDialog.showSaveDialogSync({ title: 'پاشکەوتکردن', defaultPath: path.join(os.homedir(), 'Desktop', fileName), filters: [{ name: 'Excel', extensions: ['xlsx'] }] })
            : path.join(os.homedir(), 'Desktop', fileName);
        if (savePath) { fs.writeFileSync(savePath, buf); showToast('✅ Excel پاشکەوت کرا!'); }
    } catch (err) { showToast('❌ هەڵە: ' + err.message); }
});

// Note: Payroll triggers consolidated in the main navigation listener at the top of the file

// Initial payroll data load (for stats)
updatePayrollStats();
renderQuickClockList();
if (window.updateNotifications) window.updateNotifications();

// Excel Import Logic (Keep existing)
pasteArea.addEventListener('input', () => {
    const raw = pasteArea.value.trim();
    if (!raw) { pastePreview.innerHTML = ''; return; }
    const parsed = parsePastedData(raw);
    if (parsed.length === 0) { pastePreview.innerHTML = '<p class="paste-error">❌ هەڵە لە داتا</p>'; return; }
    renderPastePreview(parsed);
});

function parsePastedData(raw) {
    const rows = raw.split('\n').filter(r => r.trim() !== '');
    return rows.map(row => {
        const cols = row.split('\t');
        if (cols.length < 6) return null;
        return {
            id: Date.now() + Math.random(),
            batchNo: cols[0]?.trim(), code: cols[1]?.trim(), name: cols[2]?.trim(),
            qty: parseInt(cols[3]?.replace(/,/g, '')) || 0,
            location: cols[4]?.trim(), stickerQty: parseInt(cols[5]?.replace(/,/g, '')) || 0,
            price: parseInt(cols[6]?.replace(/,/g, '')) || 0,
            pNumber: cols[7]?.trim(), company: cols[8]?.trim(),
            receivedDate: cols[9]?.trim(), expDate: cols[10]?.trim()
        };
    }).filter(p => p !== null);
}

function importPastedData() {
    const parsed = parsePastedData(pasteArea.value.trim());
    if (parsed.length === 0) return;
    products.push(...parsed);
    saveData('sticker_products', products);

    // Auto-save imported items to catalog
    parsed.forEach(newProduct => {
        if (!newProduct.name) return;
        const existingIndex = catalog.findIndex(item => 
            (newProduct.batchNo && item.batchNo && item.batchNo.toLowerCase() === newProduct.batchNo.toLowerCase()) || 
            (newProduct.code && item.code && item.code.toLowerCase() === newProduct.code.toLowerCase())
        );
        const templateData = {
            id: existingIndex >= 0 ? catalog[existingIndex].id : Date.now() + Math.random(),
            batchNo: newProduct.batchNo || '',
            code: newProduct.code || '',
            name: newProduct.name,
            price: newProduct.price || 0,
            location: newProduct.location || '',
            company: newProduct.company || '',
            pNumber: newProduct.pNumber || '',
            expDate: newProduct.expDate || ''
        };
        if (existingIndex >= 0) {
            catalog[existingIndex] = templateData;
        } else {
            catalog.push(templateData);
        }
    });
    saveData('sticker_catalog', catalog);

    renderTable();
    updateSalesDropdown();
    closeModal();
    showToast(`${parsed.length} دانە زیادکرا!`);
}

function renderPastePreview(data) {
    let html = `<p class="preview-count">✅ ${data.length} دانە دۆزرایەوە</p><table class="data-table"><tbody>`;
    data.slice(0, 5).forEach(p => { html += `<tr><td>${p.name}</td><td>${p.batchNo}</td></tr>`; });
    html += '</tbody></table>';
    pastePreview.innerHTML = html;
}

// ==================== EXCEL EXPORT SYSTEM ====================
let currentExportType = 'full'; // 'products' | 'sales' | 'workers' | 'report' | 'full'

const exportModal = document.getElementById('exportModal');
const closeExportModalBtn = document.getElementById('closeExportModalBtn');
const cancelExportBtn = document.getElementById('cancelExportBtn');
const doExportBtn = document.getElementById('doExportBtn');
const exportTypeLabel = document.getElementById('exportTypeLabel');
const exportWorkerFilter = document.getElementById('exportWorkerFilter');

function openExportModal(type) {
    currentExportType = type;
    exportModal.classList.add('active');
    document.getElementById('exportStatusMsg').style.display = 'none';

    // Set dates to current month by default
    const now = new Date();
    const firstDay = new Date(now.getFullYear(), now.getMonth(), 1).toISOString().split('T')[0];
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0).toISOString().split('T')[0];
    document.getElementById('exportFromDate').value = firstDay;
    document.getElementById('exportToDate').value = lastDay;

    // Populate worker dropdown
    exportWorkerFilter.innerHTML = '<option value="">هەموو کڕێکارەکان</option>';
    workers.forEach(w => {
        const opt = document.createElement('option');
        opt.value = w.id;
        opt.textContent = w.name;
        exportWorkerFilter.appendChild(opt);
    });

    // Show/hide worker filter based on type
    const workerFilterDiv = document.getElementById('exportWorkerFilterDiv');

    // Set checkboxes and labels based on type
    const chkProducts = document.getElementById('chkProducts');
    const chkSales = document.getElementById('chkSales');
    const chkReport = document.getElementById('chkWorkerReport');

    const labels = {
        products: '📦 هەناردەکردنی لیستی بەرهەمەکان',
        workers: '👷 هەناردەکردنی لیستی کڕێکارەکان',
        sales: '🛒 هەناردەکردنی مێژووی فرۆشتنەکان',
        report: '📊 هەناردەکردنی ڕاپۆرتی کڕێکارەکان',
        full: '📋 هەناردەکردنی هەموو داتاکان (فایلی تەواو)'
    };
    exportTypeLabel.textContent = labels[type] || '';

    if (type === 'products') {
        chkProducts.checked = true; chkSales.checked = false; chkReport.checked = false;
        workerFilterDiv.style.display = 'none';
    } else if (type === 'workers') {
        chkProducts.checked = false; chkSales.checked = false; chkReport.checked = true;
        workerFilterDiv.style.display = 'block';
    } else if (type === 'sales') {
        chkProducts.checked = false; chkSales.checked = true; chkReport.checked = false;
        workerFilterDiv.style.display = 'block';
    } else if (type === 'report') {
        chkProducts.checked = false; chkSales.checked = false; chkReport.checked = true;
        workerFilterDiv.style.display = 'block';
    } else {
        chkProducts.checked = true; chkSales.checked = true; chkReport.checked = true;
        workerFilterDiv.style.display = 'block';
    }
}

function closeExportModal() {
    exportModal.classList.remove('active');
}

closeExportModalBtn.addEventListener('click', closeExportModal);
cancelExportBtn.addEventListener('click', closeExportModal);

// Hook export buttons
document.getElementById('exportProductsBtn').addEventListener('click', () => openExportModal('products'));
document.getElementById('exportWorkersBtn').addEventListener('click', () => openExportModal('workers'));
document.getElementById('exportSalesBtn').addEventListener('click', () => openExportModal('sales'));
document.getElementById('exportReportBtn').addEventListener('click', () => openExportModal('report'));

doExportBtn.addEventListener('click', () => {
    const fromDate = document.getElementById('exportFromDate').value;
    const toDate = document.getElementById('exportToDate').value;
    const filterWorkerId = exportWorkerFilter.value;
    const includeProducts = document.getElementById('chkProducts').checked;
    const includeSales = document.getElementById('chkSales').checked;
    const includeReport = document.getElementById('chkWorkerReport').checked;

    if (!includeProducts && !includeSales && !includeReport) {
        showExportStatus('⚠️ تکایە لانیکەم یەک شیت هەڵبژێرە!', 'error');
        return;
    }

    const wb = XLSX.utils.book_new();

    // ---- SHEET 1: Products ----
    if (includeProducts) {
        const prodData = [
            ['ژمارەی باچ', 'کۆدی بەرهەم', 'ناوی بەرهەم', 'بڕ (QTY)', 'شوێن', 'بڕی ستیکەر', 'نرخ (IQD)', 'ژ. مۆڵەت', 'کۆمپانیا', 'بەرواری وەرگرتن', 'بەرواری بەسەرچوون']
        ];
        products.forEach(p => {
            prodData.push([
                p.batchNo, p.code, p.name, p.qty, p.location,
                p.stickerQty, p.price, p.pNumber, p.company,
                formatDate(p.receivedDate), formatDate(p.expDate)
            ]);
        });
        const wsProducts = XLSX.utils.aoa_to_sheet(prodData);
        // Style column widths
        wsProducts['!cols'] = [
            { wch: 14 }, { wch: 16 }, { wch: 35 }, { wch: 10 }, { wch: 10 },
            { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 20 }, { wch: 14 }, { wch: 14 }
        ];
        XLSX.utils.book_append_sheet(wb, wsProducts, 'بەرهەمەکان');
    }

    // ---- SHEET 2: Sales History ----
    if (includeSales) {
        let filteredSales = sales;

        if (fromDate) filteredSales = filteredSales.filter(s => new Date(s.date) >= new Date(fromDate));
        if (toDate) filteredSales = filteredSales.filter(s => new Date(s.date) <= new Date(toDate + 'T23:59:59'));
        if (filterWorkerId) {
            filteredSales = filteredSales.filter(s =>
                (s.assignedWorkers || []).some(w => w.id == filterWorkerId)
            );
        }

        const salesData = [
            ['ناوی بەرهەم', 'شوێنی مەبەست', 'بڕی QTY', 'کۆی ستیکەر', 'کڕێکارەکان و بڕی لێدراو', 'بەروار و کات']
        ];
        filteredSales.forEach(s => {
            const workersStr = (s.assignedWorkers || []).map(w => `${w.name}: ${w.qty}`).join(' | ');
            salesData.push([
                s.productName,
                s.destination || '-',
                s.qty,
                s.stickerQty,
                workersStr,
                new Date(s.date).toLocaleString('ku-IQ')
            ]);
        });

        // Add totals row
        const totalQty = filteredSales.reduce((a, s) => a + s.qty, 0);
        const totalStickers = filteredSales.reduce((a, s) => a + s.stickerQty, 0);
        salesData.push(['', '', '', '', '', '']);
        salesData.push(['کۆی گشتی', '', totalQty, totalStickers, '', '']);

        const wsSales = XLSX.utils.aoa_to_sheet(salesData);
        wsSales['!cols'] = [{ wch: 35 }, { wch: 18 }, { wch: 12 }, { wch: 14 }, { wch: 40 }, { wch: 22 }];
        XLSX.utils.book_append_sheet(wb, wsSales, 'مێژووی فرۆشتن');
    }

    // ---- SHEET 3: Worker Monthly Report ----
    if (includeReport) {
        let filteredSalesForReport = sales;
        if (fromDate) filteredSalesForReport = filteredSalesForReport.filter(s => new Date(s.date) >= new Date(fromDate));
        if (toDate) filteredSalesForReport = filteredSalesForReport.filter(s => new Date(s.date) <= new Date(toDate + 'T23:59:59'));

        const workerStats = {};
        filteredSalesForReport.forEach(sale => {
            (sale.assignedWorkers || []).forEach(w => {
                if (filterWorkerId && w.id != filterWorkerId) return;
                if (!workerStats[w.id]) {
                    workerStats[w.id] = { name: w.name, totalStickers: 0, jobsCount: 0, lastDate: sale.date, salesList: [] };
                }
                workerStats[w.id].totalStickers += w.qty;
                workerStats[w.id].jobsCount += 1;
                if (new Date(sale.date) > new Date(workerStats[w.id].lastDate)) {
                    workerStats[w.id].lastDate = sale.date;
                }
                workerStats[w.id].salesList.push({
                    product: sale.productName,
                    destination: sale.destination,
                    qty: w.qty,
                    date: sale.date
                });
            });
        });

        const reportData = [
            ['ناوی کڕێکار', 'کۆی ستیکەری لێدراو', 'ژمارەی ئیشەکان', 'دوایین چالاکی']
        ];
        const statsArr = Object.values(workerStats).sort((a, b) => b.totalStickers - a.totalStickers);
        statsArr.forEach(stat => {
            reportData.push([
                stat.name,
                stat.totalStickers,
                stat.jobsCount,
                new Date(stat.lastDate).toLocaleString('ku-IQ')
            ]);
        });

        // Grand total row
        const grandTotal = statsArr.reduce((a, s) => a + s.totalStickers, 0);
        const grandJobs = statsArr.reduce((a, s) => a + s.jobsCount, 0);
        reportData.push(['', '', '', '']);
        reportData.push(['کۆی گشتی', grandTotal, grandJobs, '']);

        const wsReport = XLSX.utils.aoa_to_sheet(reportData);
        wsReport['!cols'] = [{ wch: 25 }, { wch: 20 }, { wch: 18 }, { wch: 24 }];
        XLSX.utils.book_append_sheet(wb, wsReport, 'ڕاپۆرتی کڕێکارەکان');

        // ---- Sub-sheet: Detail per worker (if filter selected) ----
        if (filterWorkerId && workerStats[filterWorkerId]) {
            const stat = workerStats[filterWorkerId];
            const detailData = [
                [`ڕاپۆرتی تایبەتی: ${stat.name}`],
                ['بەرهەم', 'شوێنی مەبەست', 'بڕی ستیکەر', 'بەروار']
            ];
            stat.salesList.forEach(sl => {
                detailData.push([sl.product, sl.destination || '-', sl.qty, new Date(sl.date).toLocaleString('ku-IQ')]);
            });
            detailData.push(['', '', '', '']);
            detailData.push(['کۆی گشتی', '', stat.totalStickers, '']);
            const wsDetail = XLSX.utils.aoa_to_sheet(detailData);
            wsDetail['!cols'] = [{ wch: 35 }, { wch: 18 }, { wch: 16 }, { wch: 22 }];
            XLSX.utils.book_append_sheet(wb, wsDetail, `${stat.name.substring(0, 20)}`);
        }
    }

    // ---- Save file ----
    const now = new Date();
    const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    const defaultName = `StickerPro_Report_${dateStr}.xlsx`;

    try {
        const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        if (electronDialog) {
            // Show native save dialog
            const savePath = electronDialog.showSaveDialogSync({
                title: 'پاشکەوتکردنی فایلی Excel',
                defaultPath: path.join(os.homedir(), 'Desktop', defaultName),
                filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
            });
            if (savePath) {
                fs.writeFileSync(savePath, buf);
                showExportStatus(`✅ پاشکەوت کرا:\n${savePath}`, 'success');
                setTimeout(closeExportModal, 2500);
                showToast('✅ Excel بە سەرکەوتوویی هەناردەکرا!');
            }
        } else {
            // Fallback: save directly to Desktop
            const desktopPath = path.join(os.homedir(), 'Desktop', defaultName);
            fs.writeFileSync(desktopPath, buf);
            showExportStatus(`✅ فایلەکە لە Desktop پاشکەوت کرا!\n📁 ${defaultName}`, 'success');
            setTimeout(closeExportModal, 2500);
            showToast('✅ Excel لە Desktop پاشکەوت کرا!');
        }
    } catch (err) {
        console.error('Export error:', err);
        // Last resort fallback
        try {
            const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
            const fallbackPath = path.join(os.homedir(), defaultName);
            fs.writeFileSync(fallbackPath, buf);
            showExportStatus(`✅ پاشکەوت کرا: ${fallbackPath}`, 'success');
            setTimeout(closeExportModal, 2500);
            showToast('✅ Excel پاشکەوت کرا!');
        } catch (err2) {
            showExportStatus('❌ هەڵە: ' + err2.message, 'error');
        }
    }
});

function showExportStatus(msg, type) {
    const el = document.getElementById('exportStatusMsg');
    el.style.display = 'block';
    el.style.background = type === 'success' ? 'rgba(34,197,94,0.15)' : 'rgba(239,68,68,0.15)';
    el.style.color = type === 'success' ? '#22c55e' : '#ef4444';
    el.style.border = `1px solid ${type === 'success' ? '#22c55e' : '#ef4444'}`;
    el.textContent = msg;
}

// ==================== SMART NOTIFICATION PANEL LOGIC ====================
function updateNotifications() {
    const badge = document.getElementById('notificationBadge');
    const list = document.getElementById('notificationList');
    if (!badge || !list) return;

    const alerts = [];
    const now = Date.now();

    products.forEach(p => {
        // 1. Expiration check (Expired vs Expiring in 90 Days)
        if (p.expDate) {
            const expTime = new Date(p.expDate).getTime();
            const diffMs = expTime - now;
            const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));

            if (diffDays <= 0) {
                alerts.push({
                    id: `exp-zero-${p.id}`,
                    type: 'danger',
                    title: `⚠️ بەسەرچووە: ${p.name}`,
                    desc: `باچ: ${p.batchNo || '-'} | شوێن: ${p.location || '-'} (بەسەرچووە!)`,
                    action: `deleteExpiredProduct(${p.id})`,
                    actionLabel: '🗑️ سڕینەوە'
                });
            } else if (diffDays <= 90) {
                alerts.push({
                    id: `exp-warn-${p.id}`,
                    type: 'warning',
                    title: `⏳ بەسەرچوون نزیکە: ${p.name}`,
                    desc: `تەنها ${diffDays} ڕۆژ ماوە! (شوێن: ${p.location || '-'})`,
                    action: `adjustNotificationStock(${p.id})`,
                    actionLabel: '🔄 چارەسەر'
                });
            }
        }

        // 2. Low Stock check
        if (p.qty <= 0) {
            alerts.push({
                id: `stock-zero-${p.id}`,
                type: 'danger',
                title: `❌ کۆگا تەواو بووە: ${p.name}`,
                desc: `ڕەفەی ${p.location || '-'} (بڕی مەوجود گەیشتووەتە 0!)`,
                action: `adjustNotificationStock(${p.id})`,
                actionLabel: '🔄 زیادکردن'
            });
        } else if (p.qty < 5 || (p.stickerQty && p.stickerQty < 500)) {
            alerts.push({
                id: `stock-low-${p.id}`,
                type: 'warning',
                title: `⚠️ مەوجود کەمە: ${p.name}`,
                desc: `ماوە: ${p.qty} کارتۆن | ${p.stickerQty || 0} ستیکەر (ڕەفە: ${p.location || '-'})`,
                action: `adjustNotificationStock(${p.id})`,
                actionLabel: '🔄 زیادکردن'
            });
        }
    });

    // Update Badge
    if (alerts.length > 0) {
        badge.textContent = alerts.length;
        badge.style.display = 'block';
    } else {
        badge.style.display = 'none';
    }

    // Render list
    list.innerHTML = '';
    if (alerts.length === 0) {
        list.innerHTML = `
            <div style="text-align: center; color: #64748b; padding: 25px 0; font-size: 0.85rem; direction: rtl;">
                هیچ ئاگادارکردنەوەیەک نییە! کۆگا تەندروستە. 🎉
            </div>
        `;
        return;
    }

    alerts.forEach(a => {
        const item = document.createElement('div');
        item.style.cssText = `
            display: flex;
            flex-direction: column;
            gap: 6px;
            padding: 12px 16px;
            border-bottom: 1px solid rgba(0,0,0,0.04);
            background: ${a.type === 'danger' ? 'rgba(239, 68, 68, 0.03)' : 'rgba(251, 191, 36, 0.02)'};
            border-right: 4px solid ${a.type === 'danger' ? '#ef4444' : '#fbbf24'};
            transition: all 0.2s ease;
        `;

        item.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: flex-start; gap: 8px;">
                <h5 style="margin: 0; color: #1f2937; font-size: 0.82rem; font-weight: 600; text-align: right; width: 100%;">${a.title}</h5>
                <button onclick="${a.action}" style="background: rgba(0,0,0,0.05); color: #1f2937; border: 1px solid rgba(0,0,0,0.1); border-radius: 6px; padding: 3px 8px; font-size: 0.72rem; cursor: pointer; transition: all 0.2s; white-space: nowrap; margin-right: 8px;">
                    ${a.actionLabel}
                </button>
            </div>
            <p style="margin: 0; color: #94a3b8; font-size: 0.72rem; line-height: 1.4; text-align: right;">${a.desc}</p>
        `;
        list.appendChild(item);
    });
}

function adjustNotificationStock(productId) {
    const dropdown = document.getElementById('notificationDropdown');
    if (dropdown) dropdown.style.display = 'none';

    if (window.quickAdjustStock) {
        window.quickAdjustStock(productId);
    } else {
        showToast('⚠️ تکایە لە بەشی مەوجودات کارەکە بکە.');
    }
}

function deleteExpiredProduct(productId) {
    const p = products.find(prod => prod.id == productId);
    if (!p) return;

    if (confirm(`⚠️ ئایا دڵنیایت لە سڕینەوەی ئەم کاڵا بەسەرچووە؟\n📦 ناوی کاڵا: ${p.name}`)) {
        products = products.filter(prod => prod.id !== productId);
        saveData('sticker_products', products);

        renderTable();
        renderProfessionalInventory();
        updateNotifications();
        updateStats();

        showToast('🗑️ بەرهەمە بەسەرچووەکە بە سەرکەوتوویی سڕایەوە!');
    }
}

window.adjustNotificationStock = adjustNotificationStock;
window.deleteExpiredProduct = deleteExpiredProduct;
window.updateNotifications = updateNotifications;

function initNotificationListeners() {
    const bellBtn = document.getElementById('notificationBellBtn');
    const dropdown = document.getElementById('notificationDropdown');
    const closeBtn = document.getElementById('closeNotificationsBtn');

    if (bellBtn && dropdown) {
        bellBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const isOpen = dropdown.style.display === 'block';
            dropdown.style.display = isOpen ? 'none' : 'block';
            if (!isOpen) updateNotifications();
        });
    }

    if (closeBtn && dropdown) {
        closeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            dropdown.style.display = 'none';
        });
    }

    document.addEventListener('click', (e) => {
        if (dropdown && !dropdown.contains(e.target) && e.target !== bellBtn) {
            dropdown.style.display = 'none';
        }
    });
}
initNotificationListeners();

// ==================== INTERACTIVE WAREHOUSE SHELF MAP LOGIC ====================

function renderWarehouseShelfMap() {
    const grid = document.getElementById('shelfMapGrid');
    if (!grid) return;

    // Define Core shelves
    const coreShelves = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2', 'D1', 'D2'];

    // Scan products for any additional custom locations to build a truly dynamic layout!
    const customShelves = new Set();
    products.forEach(p => {
        if (p.location) {
            const cleanLoc = p.location.replace(/📍\s*شوێن\s*:\s*/g, '').trim().toUpperCase();
            if (cleanLoc && !coreShelves.includes(cleanLoc)) {
                customShelves.add(cleanLoc);
            }
        }
    });

    const allShelves = [...coreShelves, ...Array.from(customShelves)].sort();

    grid.innerHTML = '';

    allShelves.forEach(shelf => {
        // Find products on this shelf
        const shelfProducts = products.filter(p => {
            if (!p.location) return false;
            const clean = p.location.replace(/📍\s*شوێن\s*:\s*/g, '').trim().toUpperCase();
            return clean === shelf;
        });

        const totalQty = shelfProducts.reduce((sum, p) => sum + (p.qty || 0), 0);
        const totalStickers = shelfProducts.reduce((sum, p) => sum + (p.stickerQty || 0), 0);

        // Check if any product on this shelf is low stock or expired
        const hasDanger = shelfProducts.some(p => {
            const exp = new Date(p.expDate);
            const now = new Date();
            const isExpired = exp < now;
            return p.qty <= 5 || p.stickerQty < 500 || isExpired;
        });

        // Determine visual theme based on shelf state
        let cardBg = 'rgba(0,0,0,0.02)';
        let cardBorder = 'rgba(0,0,0,0.08)';
        let badgeColor = '#64748b';
        let badgeBg = 'rgba(0,0,0,0.05)';
        let badgeLabel = 'بەتاڵ 📭';

        if (shelfProducts.length > 0) {
            if (hasDanger) {
                cardBg = 'rgba(239, 68, 68, 0.05)';
                cardBorder = 'rgba(239, 68, 68, 0.25)';
                badgeColor = '#ef4444';
                badgeBg = 'rgba(239, 68, 68, 0.15)';
                badgeLabel = 'مەترسی 🔴';
            } else {
                cardBg = 'rgba(34, 197, 94, 0.03)';
                cardBorder = 'rgba(34, 197, 94, 0.2)';
                badgeColor = '#22c55e';
                badgeBg = 'rgba(34, 197, 94, 0.1)';
                badgeLabel = 'سەلامەت 🟢';
            }
        }

        const isSelected = activeSelectedShelf === shelf;
        if (isSelected) {
            cardBorder = '#a855f7';
            cardBg = 'rgba(168, 85, 247, 0.08)';
        }

        const shelfCard = document.createElement('div');
        shelfCard.style.cssText = `
            background: ${cardBg};
            border: 2px solid ${cardBorder};
            border-radius: 12px;
            padding: 12px;
            cursor: pointer;
            text-align: center;
            transition: all 0.2s ease;
            box-shadow: ${isSelected ? '0 0 12px rgba(168, 85, 247, 0.3)' : 'none'};
        `;

        shelfCard.innerHTML = `
            <div style="font-weight: 700; color: #1f2937; font-size: 1rem; margin-bottom: 4px;">ڕەفەی ${shelf}</div>
            <span style="font-size: 0.65rem; padding: 2px 8px; border-radius: 10px; color: ${badgeColor}; background: ${badgeBg}; font-weight: bold;">
                ${badgeLabel}
            </span>
            <div style="font-size: 0.72rem; color: #6b7280; margin-top: 8px; border-top: 1px solid rgba(0,0,0,0.05); padding-top: 6px;">
                📦 ${shelfProducts.length} جۆر کاڵا<br/>
                🏷️ ${totalStickers.toLocaleString()} ستیکەر
            </div>
        `;

        shelfCard.addEventListener('click', () => {
            activeSelectedShelf = shelf;
            renderWarehouseShelfMap();
            showShelfDetails(shelf, shelfProducts);
        });

        grid.appendChild(shelfCard);
    });

    // Auto-update right panel if a shelf is currently selected
    if (activeSelectedShelf) {
        const shelfProducts = products.filter(p => {
            if (!p.location) return false;
            const clean = p.location.replace(/📍\s*شوێن\s*:\s*/g, '').trim().toUpperCase();
            return clean === activeSelectedShelf;
        });
        showShelfDetails(activeSelectedShelf, shelfProducts);
    }
}

function showShelfDetails(shelf, shelfProducts) {
    const title = document.getElementById('shelfDetailTitle');
    const content = document.getElementById('shelfDetailContent');
    if (!title || !content) return;

    title.innerText = `📍 بەرهەمەکانی ڕەفەی ${shelf}`;

    if (shelfProducts.length === 0) {
        content.innerHTML = `
            <div style="text-align: center; color: #64748b; padding-top: 50px;">
                <span style="font-size: 2.5rem; display: block; margin-bottom: 10px;">📭</span>
                ئەم ڕەفەیە لە ئێستادا بەتاڵە و هیچ بەرهەمێکی لەسەر نییە!
            </div>
        `;
        return;
    }

    let html = `
        <div style="display: flex; flex-direction: column; gap: 10px; max-height: 280px; overflow-y: auto; padding-right: 5px;">
    `;

    shelfProducts.forEach(p => {
        const isLow = p.qty <= 5 || p.stickerQty < 500;
        const borderStyle = isLow ? '1px solid rgba(239, 68, 68, 0.25)' : '1px solid var(--border-color)';
        const bgStyle = isLow ? 'rgba(239, 68, 68, 0.05)' : 'var(--modal-input-bg)';

        html += `
            <div style="background: ${bgStyle}; border: ${borderStyle}; border-radius: 8px; padding: 10px; display: flex; justify-content: space-between; align-items: center; transition: all 0.2s; direction: rtl;">
                <div style="text-align: right; flex: 1; padding-left: 10px;">
                    <div style="font-weight: 600; color: var(--text-main); font-size: 0.82rem; margin-bottom: 2px;">${p.name}</div>
                    <div style="font-size: 0.7rem; color: var(--text-dim);">
                        کۆد: <code>${p.code || '-'}</code> | باچ: <code>${p.batchNo || '-'}</code>
                    </div>
                    <div style="font-size: 0.72rem; margin-top: 5px;">
                        <span style="color: ${p.qty <= 5 ? '#ef4444' : '#38bdf8'}; font-weight: bold;">📦 ${p.qty} QTY</span>
                        <span style="color: var(--text-dim); margin: 0 5px;">|</span>
                        <span style="color: ${p.stickerQty < 500 ? '#ef4444' : '#fbbf24'}; font-weight: bold;">🏷️ ${p.stickerQty.toLocaleString()} دانە</span>
                    </div>
                </div>
                <button onclick="window.quickAdjustStock(${p.id})" style="background: rgba(56, 189, 248, 0.1); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.25); border-radius: 6px; padding: 6px 10px; font-size: 0.75rem; cursor: pointer; transition: all 0.2s; white-space: nowrap; margin-right: 8px;">
                    🔄 ڕێکخستن
                </button>
            </div>
        `;
    });

    html += `</div>`;
    content.innerHTML = html;
}

window.renderWarehouseShelfMap = renderWarehouseShelfMap;
window.quickAdjustStock = quickAdjustStock;

// ==================== THEME TOGGLE LOGIC ====================
function initThemeToggle() {
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const themeToggleIcon = document.getElementById('themeToggleIcon');
    const themeToggleText = document.getElementById('themeToggleText');

    function updateToggleUI(isDark) {
        if (themeToggleIcon) themeToggleIcon.textContent = isDark ? '☀️' : '🌙';
        if (themeToggleText) themeToggleText.textContent = isDark ? 'دۆخی ڕووناک' : 'دۆخی تاریک';
    }

    // Initialize UI state
    const isDark = document.body.classList.contains('dark-theme');
    updateToggleUI(isDark);

    if (themeToggleBtn) {
        themeToggleBtn.addEventListener('click', () => {
            document.body.classList.toggle('dark-theme');
            const newIsDark = document.body.classList.contains('dark-theme');
            localStorage.setItem('theme', newIsDark ? 'dark' : 'light');
            updateToggleUI(newIsDark);
        });
    }
}

// Run theme initializer
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initThemeToggle);
} else {
    initThemeToggle();
}

// ==================== TALAF (WASTE/DAMAGE) LOGIC ====================

function updateTalafDropdown() {
    const sel = document.getElementById('talafProductSelect');
    if (!sel) return;
    sel.innerHTML = '<option value="">هەڵبژێرە...</option>';
    products.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.id;
        opt.textContent = `${p.name} (باچ: ${p.batchNo || '-'} | 📍 ${p.location || '-'})`;
        sel.appendChild(opt);
    });
}

// Show stock info when product is selected in talaf form
const talafProductSel = document.getElementById('talafProductSelect');
if (talafProductSel) {
    talafProductSel.addEventListener('change', () => {
        const pid = talafProductSel.value;
        const ind = document.getElementById('talafStockIndicator');
        const stickerHint = document.getElementById('talafStickerHint');
        const qtyHint = document.getElementById('talafQtyHint');
        if (!pid) {
            ind.style.display = 'none';
            if (stickerHint) stickerHint.style.display = 'none';
            if (qtyHint) qtyHint.style.display = 'none';
            return;
        }
        const p = products.find(pr => pr && pr.id && pr.id.toString() === pid.toString());
        if (p) {
            document.getElementById('talafCurrentStock').textContent = (p.qty || 0).toLocaleString();
            document.getElementById('talafCurrentSticker').textContent = (p.stickerQty || 0).toLocaleString();
            ind.style.display = 'flex';
            // Update hints under each input
            if (stickerHint) {
                document.getElementById('talafStickerAvail').textContent = (p.stickerQty || 0).toLocaleString();
                stickerHint.style.display = 'block';
            }
            if (qtyHint) {
                document.getElementById('talafQtyAvail').textContent = (p.qty || 0).toLocaleString();
                qtyHint.style.display = 'block';
            }
        } else {
            ind.style.display = 'none';
            if (stickerHint) stickerHint.style.display = 'none';
            if (qtyHint) qtyHint.style.display = 'none';
        }
    });
}

// Submit talaf form
const talafForm = document.getElementById('talafForm');
if (talafForm) {
    talafForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const pid = document.getElementById('talafProductSelect').value;
        const stickerLoss = parseInt(document.getElementById('talafStickerQty').value) || 0;
        const qtyLoss = parseInt(document.getElementById('talafQty').value) || 0;
        const reason = document.getElementById('talafReason').value.trim();

        if (!pid) { showToast('⚠️ تکایە بەرهەمێک هەڵبژێرە!'); return; }
        if (stickerLoss === 0 && qtyLoss === 0) {
            showToast('⚠️ تکایە لانیکەم یەک بڕ بنووسە!');
            return;
        }

        const pIdx = products.findIndex(p => p && p.id && p.id.toString() === pid.toString());
        if (pIdx === -1) { showToast('❌ بەرهەمەکە نەدۆزرایەوە!'); return; }

        const p = products[pIdx];

        // Validate stock
        if (stickerLoss > (p.stickerQty || 0)) {
            showToast('⚠️ بڕی \u062a\u06a9\u062c\u0648\u0648 زیاترە لە ستیکەری بەردەست!');
            return;
        }
        if (qtyLoss > (p.qty || 0)) {
            showToast('⚠️ بڕی \u062a\u06a9\u062c\u0648\u0648 زیاترە لە QTY-ی بەردەست!');
            return;
        }

        // Deduct from stock
        products[pIdx].stickerQty = (p.stickerQty || 0) - stickerLoss;
        products[pIdx].qty = (p.qty || 0) - qtyLoss;
        saveData('sticker_products', products);
        renderTable();
        updateSalesDropdown();
        updateTalafDropdown();

        // Record waste
        talaf.unshift({
            id: Date.now() + Math.random(),
            productId: pid,
            productName: p.name,
            stickerQty: stickerLoss,
            qty: qtyLoss,
            reason: reason || '-',
            date: new Date().toISOString()
        });
        saveData('sticker_talaf', talaf);

        renderTalafTable();
        updateTalafStats();
        updateStats();
        talafForm.reset();
        document.getElementById('talafStockIndicator').style.display = 'none';
        showToast(`🚮 \u062a\u06a9\u062c\u0648\u0648ەکە تۆمارکرا! ${stickerLoss} ستیکەر و ${qtyLoss} QTY کەمکرایەوە.`);
    });
}

function renderTalafTable() {
    const tbody = document.getElementById('talafTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';
    if (talaf.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:#6b7280; padding:30px;">هیچ تۆمارێکی \u062a\u06a9\u062c\u0648\u0648 نییە.</td></tr>';
        return;
    }
    talaf.forEach(t => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight:600;">${t.productName}</td>
            <td style="color:#ef4444; font-weight:700;">-${t.stickerQty.toLocaleString()}</td>
            <td style="color:#ef4444;">-${(t.qty || 0).toLocaleString()}</td>
            <td><span style="background:rgba(239,68,68,0.1); color:#ef4444; padding:3px 10px; border-radius:20px; font-size:0.8rem;">${t.reason || '-'}</span></td>
            <td>${new Date(t.date).toLocaleString('ku-IQ')}</td>
            <td style="text-align:center;">
                <div class="action-btns" style="justify-content:center;">
                    <button class="action-btn delete" title="سڕینەوە" onclick="deleteTalaf('${t.id}')">🗑️</button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function updateTalafStats() {
    const totalAll = talaf.reduce((sum, t) => sum + (t.stickerQty || 0), 0);
    const today = new Date().toDateString();
    const totalToday = talaf.filter(t => new Date(t.date).toDateString() === today)
                            .reduce((sum, t) => sum + (t.stickerQty || 0), 0);
    const el1 = document.getElementById('talafTotalAll');
    const el2 = document.getElementById('talafTotalToday');
    const el3 = document.getElementById('talafRecordCount');
    if (el1) el1.textContent = totalAll.toLocaleString();
    if (el2) el2.textContent = totalToday.toLocaleString();
    if (el3) el3.textContent = talaf.length.toLocaleString();
}

window.deleteTalaf = function(id) {
    const idx = talaf.findIndex(t => String(t.id) === String(id));
    if (idx === -1) return;
    if (!confirm('دڵنیایت لە سڕینەوەی ئەم تۆمارەی \u062a\u06a9\u062c\u0648\u0648ە؟\nتێبینی: بڕەکانی نادەگەڕێنێتەوە بۆ کۆگا.')) return;
    talaf.splice(idx, 1);
    saveData('sticker_talaf', talaf);
    renderTalafTable();
    updateTalafStats();
    showToast('تۆمارەکە سڕایەوە! 🗑️');
};

// Excel Export for Talaf
const exportTalafBtn = document.getElementById('exportTalafBtn');
if (exportTalafBtn) {
    exportTalafBtn.addEventListener('click', () => {
        if (talaf.length === 0) { showToast('⚠️ هیچ تۆمارێکی \u062a\u06a9\u062c\u0648\u0648 نییە!'); return; }
        const wb = XLSX.utils.book_new();
        const rows = talaf.map(t => ({
            'ناوی بەرهەم': t.productName,
            'بڕی ستیکەری \u062a\u06a9\u062c\u0648\u0648': t.stickerQty,
            'بڕی QTY \u062a\u06a9\u062c\u0648\u0648': t.qty || 0,
            'هۆکار': t.reason || '-',
            'بەروار': new Date(t.date).toLocaleString('ku-IQ')
        }));
        const ws = XLSX.utils.json_to_sheet(rows);
        ws['!dir'] = 'rtl';
        XLSX.utils.book_append_sheet(wb, ws, '\u062a\u06a9\u062c\u0648\u0648');
        const filename = `Talaf_${new Date().toISOString().slice(0,10)}.xlsx`;
        if (electronDialog) {
            electronDialog.showSaveDialog({
                title: 'هەناردەکردنی \u062a\u06a9\u062c\u0648\u0648 بۆ Excel',
                defaultPath: path.join(os.homedir(), 'Desktop', filename),
                filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
            }).then(result => {
                if (!result.canceled && result.filePath) {
                    XLSX.writeFile(wb, result.filePath);
                    showToast('✅ فایلەکە پاشەکەوتکرا!');
                }
            });
        } else {
            XLSX.writeFile(wb, path.join(os.homedir(), 'Desktop', filename));
            showToast('✅ فایلی Excel لە دێسکتۆپ پاشەکەوتکرا!');
        }
    });
}

// Initialize talaf on page load
updateTalafDropdown();
renderTalafTable();
updateTalafStats();

// ==================== SETTINGS / CATALOG LOGIC ====================
function renderCatalogTable() {
    const tableBody = document.getElementById('catalogTableBody');
    if (!tableBody) return;
    tableBody.innerHTML = '';
    
    const searchQuery = (document.getElementById('catalogSearch')?.value || '').trim().toLowerCase();
    
    const filteredCatalog = catalog.filter(item => {
        const batch = (item.batchNo || '').toLowerCase();
        const code = (item.code || '').toLowerCase();
        const name = (item.name || '').toLowerCase();
        const company = (item.company || '').toLowerCase();
        return batch.includes(searchQuery) || code.includes(searchQuery) || name.includes(searchQuery) || company.includes(searchQuery);
    });
    
    filteredCatalog.forEach(item => {
        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid var(--border-color)';
        
        // Show either Batch or Code in the first column
        const identifier = [item.batchNo, item.code].filter(Boolean).join(' / ') || 'بێ باچ یان کۆد';
        
        tr.innerHTML = `
            <td style="padding: 12px; font-size: 0.85rem; color: var(--text-main); font-weight: 500;">
                <span class="batch-badge" style="background: rgba(99, 102, 241, 0.1); color: #6366f1; padding: 2px 8px; border-radius: 4px; font-size: 0.8rem;">${identifier}</span>
            </td>
            <td style="padding: 12px; font-size: 0.85rem; color: var(--text-main); font-weight: 600;">${item.name || ''}</td>
            <td style="padding: 12px; font-size: 0.85rem; color: var(--text-muted);">${item.company || ''}</td>
            <td style="padding: 12px; text-align: center;">
                <div style="display: flex; gap: 8px; justify-content: center;">
                    <button onclick="editCatalogEntry(${item.id})" class="action-btn edit-btn" style="padding: 4px 8px; font-size: 0.8rem; border-radius: 6px; border: none; background: rgba(99, 102, 241, 0.1); color: #6366f1; cursor: pointer; transition: all 0.2s;">✏️ دەستکاری</button>
                    <button onclick="deleteCatalogEntry(${item.id})" class="action-btn delete-btn" style="padding: 4px 8px; font-size: 0.8rem; border-radius: 6px; border: none; background: rgba(239, 68, 68, 0.1); color: #ef4444; cursor: pointer; transition: all 0.2s;">🗑️ سڕینەوە</button>
                </div>
            </td>
        `;
        tableBody.appendChild(tr);
    });
}

// Bind search input listener
const catalogSearchInput = document.getElementById('catalogSearch');
if (catalogSearchInput) {
    catalogSearchInput.addEventListener('input', renderCatalogTable);
}

// Form Submission
const catalogForm = document.getElementById('catalogForm');
if (catalogForm) {
    catalogForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const idVal = document.getElementById('catalogId').value;
        const batchNo = document.getElementById('catalogBatchNo').value.trim();
        const code = document.getElementById('catalogProductCode').value.trim();
        const name = document.getElementById('catalogProductName').value.trim();
        const price = parseInt(document.getElementById('catalogPrice').value) || 0;
        const location = document.getElementById('catalogLocation').value.trim();
        const company = document.getElementById('catalogCompany').value.trim();
        const pNumber = document.getElementById('catalogPNumber').value.trim();
        const expDate = document.getElementById('catalogExpDate').value;
        
        if (!name) {
            showToast('⚠️ تکایە ناوی بەرهەم بنووسە.');
            return;
        }
        
        const entryData = {
            id: idVal ? parseFloat(idVal) : Date.now() + Math.random(),
            batchNo,
            code,
            name,
            price,
            location,
            company,
            pNumber,
            expDate
        };
        
        if (idVal) {
            // Edit existing
            const idx = catalog.findIndex(item => item.id === entryData.id);
            if (idx >= 0) {
                catalog[idx] = entryData;
                showToast('زانیارییەکە نوێکرایەوە! 📝');
            }
        } else {
            // Check for duplicates
            const dup = catalog.some(item => 
                (batchNo && item.batchNo && item.batchNo.toLowerCase() === batchNo.toLowerCase()) || 
                (code && item.code && item.code.toLowerCase() === code.toLowerCase())
            );
            if (dup) {
                showToast('⚠️ ئەم باچ یان کۆدە پێشتر تۆمارکراوە!');
                return;
            }
            catalog.push(entryData);
            showToast('زانیاری نوێ تۆمارکرا! ➕');
        }
        
        saveData('sticker_catalog', catalog);
        resetCatalogForm();
        renderCatalogTable();
    });
}

// Clear button
const clearCatalogBtn = document.getElementById('clearCatalogBtn');
if (clearCatalogBtn) {
    clearCatalogBtn.addEventListener('click', resetCatalogForm);
}

function resetCatalogForm() {
    if (catalogForm) catalogForm.reset();
    const catalogIdField = document.getElementById('catalogId');
    if (catalogIdField) catalogIdField.value = '';
    
    const formTitle = document.getElementById('catalogFormTitle');
    if (formTitle) formTitle.innerHTML = '📝 زیادکردنی زانیاری سەرەتایی';
}

function editCatalogEntry(id) {
    const item = catalog.find(x => x.id === id);
    if (!item) return;
    
    const catalogIdField = document.getElementById('catalogId');
    if (catalogIdField) catalogIdField.value = item.id;
    
    document.getElementById('catalogBatchNo').value = item.batchNo || '';
    document.getElementById('catalogProductCode').value = item.code || '';
    document.getElementById('catalogProductName').value = item.name || '';
    document.getElementById('catalogPrice').value = item.price || 0;
    document.getElementById('catalogLocation').value = item.location || '';
    document.getElementById('catalogCompany').value = item.company || '';
    document.getElementById('catalogPNumber').value = item.pNumber || '';
    document.getElementById('catalogExpDate').value = item.expDate || '';
    
    const formTitle = document.getElementById('catalogFormTitle');
    if (formTitle) formTitle.innerHTML = '✏️ دەستکاریکردنی زانیاری سەرەتایی';
    
    // Focus the first input field
    const batchNoInput = document.getElementById('catalogBatchNo');
    if (batchNoInput) {
        batchNoInput.focus();
        if (batchNoInput.select) batchNoInput.select();
    }
}

function deleteCatalogEntry(id) {
    if (confirm('ئایا دڵنیای لە سڕینەوەی ئەم زانیارییە سەرەتاییە؟')) {
        catalog = catalog.filter(x => x.id !== id);
        saveData('sticker_catalog', catalog);
        renderCatalogTable();
        showToast('زانیارییەکە سڕایەوە! 🗑️');
        
        // If we are currently editing the deleted item, reset the form
        const currentEditId = document.getElementById('catalogId')?.value;
        if (currentEditId && parseFloat(currentEditId) === id) {
            resetCatalogForm();
        }
    }
}

// Expose functions globally for table event bindings
window.editCatalogEntry = editCatalogEntry;
window.deleteCatalogEntry = deleteCatalogEntry;

function switchCatalogTab(tab) {
    const tabManual = document.getElementById('catalogTabManual');
    const tabPaste = document.getElementById('catalogTabPaste');
    const contentManual = document.getElementById('catalogContentManual');
    const contentPaste = document.getElementById('catalogContentPaste');
    
    if (tab === 'manual') {
        tabManual.classList.add('active');
        tabManual.style.borderBottom = '2px solid #6366f1';
        tabManual.style.color = 'var(--text-main)';
        
        tabPaste.classList.remove('active');
        tabPaste.style.borderBottom = '2px solid transparent';
        tabPaste.style.color = 'var(--text-muted)';
        
        contentManual.style.display = 'block';
        contentPaste.style.display = 'none';
    } else {
        tabPaste.classList.add('active');
        tabPaste.style.borderBottom = '2px solid #6366f1';
        tabPaste.style.color = 'var(--text-main)';
        
        tabManual.classList.remove('active');
        tabManual.style.borderBottom = '2px solid transparent';
        tabManual.style.color = 'var(--text-muted)';
        
        contentManual.style.display = 'none';
        contentPaste.style.display = 'block';
        
        const pasteArea = document.getElementById('catalogPasteArea');
        if (pasteArea) setTimeout(() => pasteArea.focus(), 100);
    }
}

function parseCatalogPastedData(raw) {
    const rows = raw.split('\n').filter(r => r.trim() !== '');
    return rows.map(row => {
        const cols = row.split('\t');
        if (cols.length < 3) return null; // Needs at least Batch, Code, and Name
        return {
            id: Date.now() + Math.random(),
            batchNo: cols[0]?.trim() || '',
            code: cols[1]?.trim() || '',
            name: cols[2]?.trim() || '',
            price: parseInt(cols[6]?.replace(/,/g, '')) || 0,
            location: cols[4]?.trim() || '',
            company: cols[8]?.trim() || '',
            pNumber: cols[7]?.trim() || '',
            expDate: cols[10]?.trim() || ''
        };
    }).filter(p => p !== null && p.name);
}

// Paste preview listener
const catalogPasteArea = document.getElementById('catalogPasteArea');
const catalogPastePreview = document.getElementById('catalogPastePreview');
if (catalogPasteArea && catalogPastePreview) {
    catalogPasteArea.addEventListener('input', () => {
        const raw = catalogPasteArea.value.trim();
        if (!raw) { catalogPastePreview.innerHTML = ''; return; }
        const parsed = parseCatalogPastedData(raw);
        if (parsed.length === 0) { catalogPastePreview.innerHTML = '<p style="color:#ef4444; font-size:0.8rem;">❌ هەڵە لە شێوازی داتا</p>'; return; }
        
        let html = `<p style="color:#10b981; font-weight:600; font-size:0.85rem; margin-bottom:8px;">✅ ${parsed.length} دۆزرایەوە (ئامادەیە بۆ هاوردەکردن)</p>`;
        html += `<table style="width:100%; font-size:0.8rem; border-collapse:collapse; border:1px solid var(--border-color); border-radius:6px;"><tbody>`;
        parsed.slice(0, 3).forEach(x => {
            html += `<tr style="border-bottom:1px solid var(--border-color);"><td style="padding:6px; color:var(--text-main); font-weight:600;">${x.name}</td><td style="padding:6px; color:var(--text-muted); text-align:left;">${x.batchNo || x.code || ''}</td></tr>`;
        });
        html += `</tbody></table>`;
        catalogPastePreview.innerHTML = html;
    });
}

// Import button listener
const importCatalogPasteBtn = document.getElementById('importCatalogPasteBtn');
if (importCatalogPasteBtn) {
    importCatalogPasteBtn.addEventListener('click', () => {
        const pasteVal = (document.getElementById('catalogPasteArea')?.value || '').trim();
        if (!pasteVal) return;
        const parsed = parseCatalogPastedData(pasteVal);
        if (parsed.length === 0) return;
        
        let addedCount = 0;
        let updatedCount = 0;
        
        parsed.forEach(newProduct => {
            const existingIndex = catalog.findIndex(item => 
                (newProduct.batchNo && item.batchNo && item.batchNo.toLowerCase() === newProduct.batchNo.toLowerCase()) || 
                (newProduct.code && item.code && item.code.toLowerCase() === newProduct.code.toLowerCase())
            );
            const templateData = {
                id: existingIndex >= 0 ? catalog[existingIndex].id : Date.now() + Math.random(),
                batchNo: newProduct.batchNo,
                code: newProduct.code,
                name: newProduct.name,
                price: newProduct.price,
                location: newProduct.location,
                company: newProduct.company,
                pNumber: newProduct.pNumber,
                expDate: newProduct.expDate
            };
            if (existingIndex >= 0) {
                catalog[existingIndex] = templateData;
                updatedCount++;
            } else {
                catalog.push(templateData);
                addedCount++;
            }
        });
        
        saveData('sticker_catalog', catalog);
        
        // Reset paste form
        if (catalogPasteArea) catalogPasteArea.value = '';
        if (catalogPastePreview) catalogPastePreview.innerHTML = '';
        switchCatalogTab('manual');
        
        renderCatalogTable();
        showToast(`✅ هاوردەکردنی بە سەرکەوتوویی تەواوبوو! (${addedCount} زیادکرا، ${updatedCount} نوێکرایەوە)`);
    });
}

window.switchCatalogTab = switchCatalogTab;

window.switchAddProductTab = function(tab) {
    const tabManual = document.getElementById('apTabManual');
    const tabPaste = document.getElementById('apTabPaste');
    const contentManual = document.getElementById('apContentManual');
    const contentPaste = document.getElementById('apContentPaste');
    
    if (tab === 'manual') {
        tabManual.classList.add('active');
        tabManual.style.borderBottom = '3px solid #6366f1';
        tabManual.style.color = 'var(--text-main)';
        
        tabPaste.classList.remove('active');
        tabPaste.style.borderBottom = '3px solid transparent';
        tabPaste.style.color = 'var(--text-muted)';
        
        contentManual.style.display = 'block';
        contentPaste.style.display = 'none';
    } else {
        tabManual.classList.remove('active');
        tabManual.style.borderBottom = '3px solid transparent';
        tabManual.style.color = 'var(--text-muted)';
        
        tabPaste.classList.add('active');
        tabPaste.style.borderBottom = '3px solid #6366f1';
        tabPaste.style.color = 'var(--text-main)';
        
        contentManual.style.display = 'none';
        contentPaste.style.display = 'block';
        
        const pasteArea = document.getElementById('apPasteArea');
        if (pasteArea) setTimeout(() => pasteArea.focus(), 100);
    }
};

const apImportPasteBtn = document.getElementById('apImportPasteBtn');
if(apImportPasteBtn) {
    apImportPasteBtn.addEventListener('click', importPastedData);
}

// ==================== CAMERA & OCR LOGIC ====================
let apCameraStream = null;

const apStartCameraBtn = document.getElementById('apStartCameraBtn');
const apCloseCameraBtn = document.getElementById('apCloseCameraBtn');
const apCaptureBtn = document.getElementById('apCaptureBtn');
const apCameraContainer = document.getElementById('apCameraScannerContainer');
const apVideoEl = document.getElementById('apWebcamVideo');
const apCanvasEl = document.getElementById('apWebcamCanvas');
const apImageInput = document.getElementById('apLabelImageInput');

if(apStartCameraBtn) {
    apStartCameraBtn.addEventListener('click', async () => {
        if(apCameraContainer) apCameraContainer.style.display = 'flex';
        try {
            apCameraStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
            if(apVideoEl) apVideoEl.srcObject = apCameraStream;
        } catch (err) {
            if(typeof showToast === 'function') showToast('⚠️ نەتوانرا کامێرا بکرێتەوە: ' + err.message);
        }
    });
}

if(apCloseCameraBtn) {
    apCloseCameraBtn.addEventListener('click', () => {
        if(apCameraStream) {
            apCameraStream.getTracks().forEach(track => track.stop());
            apCameraStream = null;
        }
        if(apCameraContainer) apCameraContainer.style.display = 'none';
    });
}

if(apCaptureBtn) {
    apCaptureBtn.addEventListener('click', () => {
        if(!apCameraStream) return;
        if(apCanvasEl && apVideoEl) {
            apCanvasEl.width = apVideoEl.videoWidth || 640;
            apCanvasEl.height = apVideoEl.videoHeight || 480;
            const ctx = apCanvasEl.getContext('2d');
            ctx.drawImage(apVideoEl, 0, 0, apCanvasEl.width, apCanvasEl.height);
            const dataUrl = apCanvasEl.toDataURL('image/jpeg');
            
            // Stop camera
            apCameraStream.getTracks().forEach(track => track.stop());
            apCameraStream = null;
            if(apCameraContainer) apCameraContainer.style.display = 'none';
            
            processOCR(dataUrl);
        }
    });
}

if(apImageInput) {
    apImageInput.addEventListener('change', (e) => {
        if(e.target.files && e.target.files[0]) {
            const reader = new FileReader();
            reader.onload = function(evt) {
                processOCR(evt.target.result);
            };
            reader.readAsDataURL(e.target.files[0]);
            e.target.value = ''; // Reset
        }
    });
}

function processOCR(imageSrc) {
    if(typeof showToast === 'function') showToast('⏳ تکایە چاوەڕێبە... سەرقاڵی خوێندنەوەی وێنەکەین');
    if(typeof Tesseract === 'undefined') {
        if(typeof showToast === 'function') showToast('❌ کتێبخانەی Tesseract بوونی نییە');
        return;
    }
    
    Tesseract.recognize(
        imageSrc,
        'eng+ara',
        { logger: m => console.log(m) }
    ).then(({ data: { text } }) => {
        console.log("OCR Result: ", text);
        if(typeof showToast === 'function') showToast('✅ وێنەکە بە سەرکەوتوویی خوێندرایەوە!');
        
        // Basic parser for demonstration
        const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
        
        let batchFound = false;
        let codeFound = false;
        let nameFound = false;
        
        lines.forEach(line => {
            const matchNum = line.match(/\d+/);
            // Attempt to assign string without numbers to Product Name
            if (!line.match(/\d/) && !nameFound && line.length > 3) {
                const el = document.getElementById('apProductName');
                if(el && !el.value) {
                    el.value = line;
                    nameFound = true;
                }
            }
            if(matchNum) {
                if(!batchFound) {
                    const el = document.getElementById('apBatchNo');
                    if(el && !el.value) {
                        el.value = matchNum[0];
                        batchFound = true;
                    }
                } else if (!codeFound) {
                    const el = document.getElementById('apProductCode');
                    if(el && !el.value) {
                        el.value = matchNum[0];
                        codeFound = true;
                    }
                }
            }
        });
    }).catch(err => {
        console.error(err);
        if(typeof showToast === 'function') showToast('❌ هەڵە لە خوێندنەوەی وێنەکە: ' + err.message);
    });
}

// ==================== AUTO FILL LOGIC ====================
function autoFillProductData(val, type) {
    if (!val || val.trim() === '') return;
    val = val.trim().toLowerCase();
    
    // Search in catalog first, then in products
    let found = catalog.find(item => item[type] && item[type].toString().toLowerCase() === val);
    if (!found) {
        found = products.find(item => item[type] && item[type].toString().toLowerCase() === val);
    }
    
    if (found) {
        let filledCount = 0;
        
        const nameEl = document.getElementById('apProductName');
        if (nameEl && !nameEl.value && found.name) { nameEl.value = found.name; filledCount++; }
        
        const priceEl = document.getElementById('apPrice');
        if (priceEl && !priceEl.value && found.price) { priceEl.value = found.price; filledCount++; }
        
        const companyEl = document.getElementById('apCompany');
        if (companyEl && !companyEl.value && found.company) { companyEl.value = found.company; filledCount++; }
        
        const pNumberEl = document.getElementById('apPNumber');
        if (pNumberEl && !pNumberEl.value && found.pNumber) { pNumberEl.value = found.pNumber; filledCount++; }
        
        const locationEl = document.getElementById('apLocation');
        if (locationEl && !locationEl.value && found.location) { locationEl.value = found.location; filledCount++; }
        
        // Also fill the other identifier if it's empty
        if (type === 'batchNo') {
            const codeEl = document.getElementById('apProductCode');
            if (codeEl && !codeEl.value && found.code) { codeEl.value = found.code; filledCount++; }
        } else if (type === 'code') {
            const batchEl = document.getElementById('apBatchNo');
            if (batchEl && !batchEl.value && found.batchNo) { batchEl.value = found.batchNo; filledCount++; }
        }
        
        if (filledCount > 0 && typeof showToast === 'function') {
            showToast('✨ زانیارییەکان بە ئۆتۆماتیکی پڕکرانەوە بەپێی داتابەیس!');
        }
    }
}

const apBatchInput = document.getElementById('apBatchNo');
if (apBatchInput) {
    apBatchInput.addEventListener('input', (e) => {
        autoFillProductData(e.target.value, 'batchNo');
    });
}

const apCodeInput = document.getElementById('apProductCode');
if (apCodeInput) {
    apCodeInput.addEventListener('input', (e) => {
        autoFillProductData(e.target.value, 'code');
    });
}
