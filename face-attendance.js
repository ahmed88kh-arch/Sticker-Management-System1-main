// ==================== FACE ATTENDANCE (AI) LOGIC ====================

let faceModelsLoaded = false;
let faceCameraStream = null;
let isFaceScannerActive = false;
let isEnrollmentMode = false;
let faceMatcher = null;
let faceDetectionInterval = null;

const MODEL_URL = 'https://cdn.jsdelivr.net/npm/@vladmandic/face-api/model/';

// Helper to add timeout to promises
const withTimeout = (promise, ms) => {
    let timeoutId;
    const timeoutPromise = new Promise((_, reject) => {
        timeoutId = setTimeout(() => reject(new Error('کێشە لە هێڵی ئینتەرنێت هەیە (Timeout)')), ms);
    });
    return Promise.race([promise, timeoutPromise]).finally(() => clearTimeout(timeoutId));
};

// DOM Elements
const faceEnrollWorkerSelect = document.getElementById('faceEnrollWorkerSelect');
const startFaceEnrollBtn = document.getElementById('startFaceEnrollBtn');
const captureEnrollFaceBtn = document.getElementById('captureEnrollFaceBtn');
const enrollStatusMsg = document.getElementById('enrollStatusMsg');

const startFaceScannerBtn = document.getElementById('startFaceScannerBtn');
const scannerStatusMsg = document.getElementById('scannerStatusMsg');

const faceCameraContainer = document.getElementById('faceCameraContainer');
const faceWebcamVideo = document.getElementById('faceWebcamVideo');
const faceWebcamCanvas = document.getElementById('faceWebcamCanvas');
const closeFaceCameraBtn = document.getElementById('closeFaceCameraBtn');

const faceLogsTableBody = document.getElementById('faceLogsTableBody');

// Load Models
async function loadFaceAPIModels() {
    if (faceModelsLoaded) return true;
    try {
        if(enrollStatusMsg) enrollStatusMsg.innerText = "⏳ سەرقاڵی دابەزاندنی مۆدێلەکان... (پێویستی بە کاتە بۆ جاری یەکەم)";
        if(scannerStatusMsg) scannerStatusMsg.innerText = "⏳ سەرقاڵی دابەزاندنی مۆدێلەکان... (پێویستی بە کاتە بۆ جاری یەکەم)";
        
        // Increased timeout to 5 minutes for slow connections in Iraq
        await withTimeout(faceapi.nets.tinyFaceDetector.loadFromUri(MODEL_URL), 300000);
        await withTimeout(faceapi.nets.faceLandmark68Net.loadFromUri(MODEL_URL), 300000);
        await withTimeout(faceapi.nets.faceRecognitionNet.loadFromUri(MODEL_URL), 300000);
        
        faceModelsLoaded = true;
        if(enrollStatusMsg) enrollStatusMsg.innerText = "✅ مۆدێلەکان دابەزین. ئێستا دەتوانیت کڕێکارەکان بناسێنیت.";
        if(scannerStatusMsg) scannerStatusMsg.innerText = "✅ سیستەم ئامادەیە بۆ کارکردن.";
        buildFaceMatcher();
        return true;
    } catch (err) {
        console.error("Face API Load Error:", err);
        if(enrollStatusMsg) enrollStatusMsg.innerText = "❌ کێشە هەیە: ئینتەرنێتەکەت زۆر خاوە یان پچڕاوە.";
        if(scannerStatusMsg) scannerStatusMsg.innerText = "❌ کێشە هەیە: ئینتەرنێتەکەت زۆر خاوە یان پچڕاوە.";
        return false;
    }
}

// Populate Worker Select
function populateFaceEnrollDropdown() {
    if(!faceEnrollWorkerSelect) return;
    faceEnrollWorkerSelect.innerHTML = '<option value="">هەڵبژێرە...</option>';
    workers.forEach(w => {
        const hasFace = w.faceDescriptor && w.faceDescriptor.length > 0;
        const opt = document.createElement('option');
        opt.value = w.id;
        opt.textContent = w.name + (hasFace ? ' (تۆمارکراوە ✅)' : '');
        faceEnrollWorkerSelect.appendChild(opt);
    });
}

// Build Face Matcher from workers
function buildFaceMatcher() {
    if (!faceModelsLoaded) return;
    const labeledDescriptors = [];
    workers.forEach(w => {
        if (w.faceDescriptor && w.faceDescriptor.length > 0) {
            const arr = new Float32Array(w.faceDescriptor);
            labeledDescriptors.push(new faceapi.LabeledFaceDescriptors(String(w.id), [arr]));
        }
    });
    
    if (labeledDescriptors.length > 0) {
        faceMatcher = new faceapi.FaceMatcher(labeledDescriptors, 0.45); // 0.45 threshold is strict enough
    } else {
        faceMatcher = null;
    }
}

// Start Camera
async function startFaceCamera(mode) {
    if (!faceModelsLoaded) {
        const loaded = await loadFaceAPIModels();
        if (!loaded) return;
    }
    
    isEnrollmentMode = (mode === 'enroll');
    isFaceScannerActive = (mode === 'scanner_in' || mode === 'scanner_out');
    
    try {
        faceCameraStream = await navigator.mediaDevices.getUserMedia({ video: { width: 640, height: 480 } });
        faceWebcamVideo.srcObject = faceCameraStream;
        faceCameraContainer.style.display = 'block';
        
        if (isEnrollmentMode) {
            captureEnrollFaceBtn.style.display = 'inline-block';
            if(enrollStatusMsg) enrollStatusMsg.innerText = "📸 ڕووخسارت بگرە بەرامبەر کامێرا و کلیک لە وێنەگرتن بکە.";
            if(scannerStatusMsg) scannerStatusMsg.innerText = "کامێرا داخراوە.";
        } else {
            captureEnrollFaceBtn.style.display = 'none';
            if(scannerStatusMsg) scannerStatusMsg.innerText = mode === 'scanner_in' ? "🟢 کامێرای هاتن چالاکە. بوەستە بەرامبەر کامێرا..." : "🔴 کامێرای چوونەوە چالاکە. بوەستە بەرامبەر کامێرا...";
            if(enrollStatusMsg) enrollStatusMsg.innerText = "کامێرا داخراوە.";
            
            faceWebcamVideo.onplay = () => {
                const displaySize = { width: faceWebcamVideo.videoWidth, height: faceWebcamVideo.videoHeight };
                faceapi.matchDimensions(faceWebcamCanvas, displaySize);
                
                faceDetectionInterval = setInterval(async () => {
                    if (!isFaceScannerActive) return;
                    
                    const detection = await faceapi.detectSingleFace(faceWebcamVideo, new faceapi.TinyFaceDetectorOptions()).withFaceLandmarks().withFaceDescriptor();
                    if (detection) {
                        const resizedDetections = faceapi.resizeResults(detection, displaySize);
                        const ctx = faceWebcamCanvas.getContext('2d');
                        ctx.clearRect(0, 0, faceWebcamCanvas.width, faceWebcamCanvas.height);
                        faceapi.draw.drawDetections(faceWebcamCanvas, resizedDetections);
                        
                        if (faceMatcher) {
                            const match = faceMatcher.findBestMatch(detection.descriptor);
                            if (match.label !== 'unknown' && match.distance < 0.45) {
                                handleFaceClockIn(match.label, mode);
                                // Pause scanner for 5 seconds to prevent spam
                                isFaceScannerActive = false;
                                setTimeout(() => { isFaceScannerActive = true; }, 5000);
                            } else {
                                scannerStatusMsg.innerText = "❌ نەناسراوە! نزیک بەرەوە یان خۆت تۆمار بکە.";
                            }
                        } else {
                            scannerStatusMsg.innerText = "⚠️ هیچ کڕێکارێک تۆمار نەکراوە بۆ ناسینەوە.";
                        }
                    } else {
                        const ctx = faceWebcamCanvas.getContext('2d');
                        ctx.clearRect(0, 0, faceWebcamCanvas.width, faceWebcamCanvas.height);
                        scannerStatusMsg.innerText = "👁️ بە دوای ڕووخساردا دەگەڕێت...";
                    }
                }, 500); // Check every 500ms
            };
        }
    } catch (err) {
        console.error(err);
        showToast('❌ نەتوانرا کامێرا بکرێتەوە!');
    }
}

// Stop Camera
function stopFaceCamera() {
    if (faceCameraStream) {
        faceCameraStream.getTracks().forEach(track => track.stop());
        faceCameraStream = null;
    }
    isFaceScannerActive = false;
    isEnrollmentMode = false;
    if (faceDetectionInterval) clearInterval(faceDetectionInterval);
    faceWebcamVideo.srcObject = null;
    faceCameraContainer.style.display = 'none';
    const ctx = faceWebcamCanvas.getContext('2d');
    if (ctx) ctx.clearRect(0, 0, faceWebcamCanvas.width, faceWebcamCanvas.height);
}

// Event Listeners
if(closeFaceCameraBtn) closeFaceCameraBtn.addEventListener('click', stopFaceCamera);

if(startFaceEnrollBtn) {
    startFaceEnrollBtn.addEventListener('click', () => {
        if (!faceEnrollWorkerSelect.value) {
            showToast('⚠️ تکایە سەرەتا کڕێکارێک هەڵبژێرە!');
            return;
        }
        startFaceCamera('enroll');
    });
}

const startFaceScannerInBtn = document.getElementById('startFaceScannerInBtn');
if(startFaceScannerInBtn) {
    startFaceScannerInBtn.addEventListener('click', () => {
        startFaceCamera('scanner_in');
    });
}

const startFaceScannerOutBtn = document.getElementById('startFaceScannerOutBtn');
if(startFaceScannerOutBtn) {
    startFaceScannerOutBtn.addEventListener('click', () => {
        startFaceCamera('scanner_out');
    });
}

if(captureEnrollFaceBtn) {
    captureEnrollFaceBtn.addEventListener('click', async () => {
        const workerId = faceEnrollWorkerSelect.value;
        if (!workerId) return;
        
        enrollStatusMsg.innerText = "⏳ خەریکی ناسینەوەی ڕووخسار...";
        const detection = await faceapi.detectSingleFace(faceWebcamVideo, new faceapi.TinyFaceDetectorOptions()).withFaceLandmarks().withFaceDescriptor();
        
        if (detection) {
            const wIdx = workers.findIndex(w => String(w.id) === String(workerId));
            if (wIdx >= 0) {
                // Save descriptor as normal Array so it can be stringified to JSON
                workers[wIdx].faceDescriptor = Array.from(detection.descriptor);
                saveData('sticker_workers', workers);
                
                showToast(`✅ دەموچاوی ${workers[wIdx].name} بە سەرکەوتوویی تۆمارکرا!`);
                enrollStatusMsg.innerText = "✅ تۆمارکرا!";
                populateFaceEnrollDropdown();
                buildFaceMatcher();
                stopFaceCamera();
            }
        } else {
            enrollStatusMsg.innerText = "❌ هیچ ڕووخسارێک نەدۆزرایەوە. تکایە ڕاست سەیر بکە.";
            showToast('❌ ڕووخسار نەدۆزرایەوە! ڕووناکی پێویست بەکاربهێنە.');
        }
    });
}

// Handle Clock In/Out from Face
function handleFaceClockIn(workerIdStr, mode) {
    const worker = workers.find(w => String(w.id) === workerIdStr);
    if (!worker) return;
    
    // Play success sound if needed, for now just toast
    
    const now = new Date();
    const timeStr = now.toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit' });
    
    // Check if currently clocked in
    const isClockedIn = !!clockInLogs[worker.id];
    
    if (mode === 'scanner_in') {
        if (!isClockedIn) {
            clockInWorker(worker.id); // Re-use existing payroll logic
            scannerStatusMsg.innerHTML = `<span style="color:#22c55e; font-size: 1.1rem; font-weight: bold;">✅ ${worker.name} دەوامی دەستپێکرد! (${timeStr})</span>`;
        } else {
            scannerStatusMsg.innerHTML = `<span style="color:#38bdf8; font-size: 1.1rem; font-weight: bold;">ℹ️ ${worker.name} پێشتر دەوامی دەستپێکردووە.</span>`;
        }
    } else if (mode === 'scanner_out') {
        if (isClockedIn) {
            clockOutWorker(worker.id); // Re-use existing payroll logic
            scannerStatusMsg.innerHTML = `<span style="color:#f59e0b; font-size: 1.1rem; font-weight: bold;">👋 ${worker.name} دەوامی تەواوکرد! (${timeStr})</span>`;
        } else {
            scannerStatusMsg.innerHTML = `<span style="color:#ef4444; font-size: 1.1rem; font-weight: bold;">⚠️ ${worker.name} هێشتا دەوامی دەستپێنەکردووە تاوەکو تەواوی بکات!</span>`;
        }
    }
    
    renderFaceLogsTable();
}

// Face Logs Rendering
function renderFaceLogsTable() {
    if (!faceLogsTableBody) return;
    faceLogsTableBody.innerHTML = '';
    
    // Get today's attendance logs for UI
    const todayStr = new Date().toISOString().split('T')[0];
    const todayLogs = attendance.filter(r => r.date === todayStr);
    
    if (todayLogs.length === 0 && Object.keys(clockInLogs).length === 0) {
        faceLogsTableBody.innerHTML = '<tr><td colspan="4" style="text-align:center;">هیچ تۆمارێک نییە</td></tr>';
        return;
    }
    
    // First, show active clock ins
    Object.keys(clockInLogs).forEach(wId => {
        const w = workers.find(x => String(x.id) === String(wId));
        if (w) {
            const timeIn = new Date(clockInLogs[wId]).toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit' });
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${w.name}</td>
                <td style="color: #22c55e; font-weight: bold;">${timeIn}</td>
                <td>-</td>
                <td>ئەمڕۆ</td>
            `;
            faceLogsTableBody.appendChild(tr);
        }
    });
    
    // Then show completed logs for today
    todayLogs.forEach(log => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${log.workerName}</td>
            <td>-</td>
            <td style="color: #f59e0b; font-weight: bold;">کۆتایی</td>
            <td>ئەمڕۆ</td>
        `;
        faceLogsTableBody.appendChild(tr);
    });
}

// Hook into existing navigation
const faceNavObserver = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
        if (mutation.target.id === 'sectionFaceAttendance' && mutation.target.style.display === 'block') {
            populateFaceEnrollDropdown();
            renderFaceLogsTable();
            buildFaceMatcher(); // rebuild just in case
        } else if (mutation.target.id === 'sectionFaceAttendance' && mutation.target.style.display === 'none') {
            stopFaceCamera();
        }
    });
});

const sectionFaceAtt = document.getElementById('sectionFaceAttendance');
if (sectionFaceAtt) {
    faceNavObserver.observe(sectionFaceAtt, { attributes: true, attributeFilter: ['style'] });
}

// ==================== INTEGRATED WORKER MODAL FACE CAPTURE ====================

let tempWorkerFaceDescriptor = null;
let workerModalStream = null;

const openWorkerFaceCameraBtn = document.getElementById('openWorkerFaceCameraBtn');
const workerModalCameraBox = document.getElementById('workerModalCameraBox');
const workerModalVideo = document.getElementById('workerModalVideo');
const captureWorkerModalFaceBtn = document.getElementById('captureWorkerModalFaceBtn');
const closeWorkerModalCameraBtn = document.getElementById('closeWorkerModalCameraBtn');
const workerModalFaceStatus = document.getElementById('workerModalFaceStatus');
const cancelWorkerBtnForm = document.getElementById('cancelWorkerBtn');
const closeWorkerModalBtnTop = document.getElementById('closeWorkerModalBtn');

async function openWorkerFaceCamera() {
    workerModalCameraBox.style.display = 'block';
    workerModalFaceStatus.innerText = "⏳ کامێرا دەکرێتەوە...";
    
    // 1. Start Camera First (Instant feedback)
    try {
        if (!workerModalStream) {
            workerModalStream = await navigator.mediaDevices.getUserMedia({ video: { width: { ideal: 640 }, height: { ideal: 480 } } });
            workerModalVideo.srcObject = workerModalStream;
        }
        workerModalFaceStatus.innerText = "📸 کامێرا کرایەوە! تکایە چاوەڕێبە تا سیستەمی AI ئامادە دەبێت...";
    } catch(err) {
        workerModalFaceStatus.innerText = "❌ کامێرا نەکرایەوە! دڵنیابە کە مۆڵەتی کامێرات داوە.";
        return; // If camera fails, stop here.
    }

    // 2. Load Models in background while camera is on
    if (!faceModelsLoaded) {
        workerModalFaceStatus.innerText = "⏳ کامێرا ئیش دەکات، سەرقاڵی دابەزاندنی مۆدێلەکان... (تکایە چاوەڕێبە)";
        const loaded = await loadFaceAPIModels();
        if (!loaded) {
            workerModalFaceStatus.innerText = "❌ کێشە لە دابەزاندنی مۆدێلەکان هەیە (ئینتەرنێت خاوە).";
            return;
        }
    }
    
    workerModalFaceStatus.innerText = "✅ ئامادەیە! ئێستا کلیک لە وێنەگرتن بکە.";
}

function stopWorkerFaceCamera() {
    if (workerModalStream) {
        workerModalStream.getTracks().forEach(t => t.stop());
        workerModalStream = null;
    }
    workerModalVideo.srcObject = null;
    if(workerModalCameraBox) workerModalCameraBox.style.display = 'none';
}

if(openWorkerFaceCameraBtn) openWorkerFaceCameraBtn.addEventListener('click', openWorkerFaceCamera);
if(closeWorkerModalCameraBtn) closeWorkerModalCameraBtn.addEventListener('click', () => {
    stopWorkerFaceCamera();
    if(workerModalFaceStatus) {
        workerModalFaceStatus.innerText = tempWorkerFaceDescriptor ? "✅ ڕووخسار ئامادەیە. فۆرمەکە خەزن بکە." : "هیچ ڕووخسارێک تۆمار نەکراوە.";
    }
});

if(captureWorkerModalFaceBtn) {
    captureWorkerModalFaceBtn.addEventListener('click', async () => {
        if (!faceModelsLoaded) {
            workerModalFaceStatus.innerText = "⚠️ تکایە چاوەڕێبە تا مۆدێلەکان بە تەواوی دادەبەزن...";
            return;
        }
        
        workerModalFaceStatus.innerText = "⏳ خەریکی ناسینەوەی ڕووخسار...";
        const detection = await faceapi.detectSingleFace(workerModalVideo, new faceapi.TinyFaceDetectorOptions()).withFaceLandmarks().withFaceDescriptor();
        
        if (detection) {
            tempWorkerFaceDescriptor = Array.from(detection.descriptor);
            workerModalFaceStatus.innerText = "✅ ڕووخسار بە سەرکەوتوویی گیرا! دەتوانیت فۆرمەکە پاشکەوت بکەیت.";
            workerModalFaceStatus.style.color = '#10b981';
            stopWorkerFaceCamera();
        } else {
            workerModalFaceStatus.innerText = "❌ نەتوانرا ڕووخسار بناسرێتەوە. ڕووناکی باشتر بکە.";
            workerModalFaceStatus.style.color = '#ef4444';
        }
    });
}

function clearModalFaceState() {
    tempWorkerFaceDescriptor = null;
    if(workerModalFaceStatus) {
        workerModalFaceStatus.innerText = "هیچ ڕووخسارێک تۆمار نەکراوە بۆ ئەم کڕێکارە.";
        workerModalFaceStatus.style.color = '#6b7280';
    }
    stopWorkerFaceCamera();
}

// Hook into the cancel/close buttons to clear temp descriptor
if(cancelWorkerBtnForm) cancelWorkerBtnForm.addEventListener('click', clearModalFaceState);
if(closeWorkerModalBtnTop) closeWorkerModalBtnTop.addEventListener('click', clearModalFaceState);

// Proxy saveData to attach face descriptor during worker save
if (typeof window.saveData === 'function') {
    const originalSaveData = window.saveData;
    window.saveData = function(key, data) {
        if (key === 'sticker_workers' && tempWorkerFaceDescriptor) {
            const wName = document.getElementById('workerName').value.trim();
            // In dashboard.js, the new worker is prepended/appended to the data array
            // We find the worker matching the form's name (which is exactly what is saved)
            const targetWorker = data.find(w => w.name === wName);
            if (targetWorker) {
                targetWorker.faceDescriptor = tempWorkerFaceDescriptor;
                clearModalFaceState(); // Consume it
                buildFaceMatcher(); // Rebuild matcher immediately
                populateFaceEnrollDropdown(); // Update the other dropdown
            }
        }
        originalSaveData(key, data);
    };
}

