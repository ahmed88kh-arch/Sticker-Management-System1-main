const loginForm = document.getElementById('loginForm');
const statusMsg = document.getElementById('statusMsg');

// Clear existing auth on login page load
localStorage.removeItem('sticker_auth');

loginForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const username = document.getElementById('username').value.trim().toLowerCase();
    const password = document.getElementById('password').value.trim();
    
    statusMsg.innerText = 'چێککردن...';
    statusMsg.className = 'status-msg';
    statusMsg.style.color = '#6b7280';

    setTimeout(() => {
        if (username === 'admin' && password === 'admin') {
            statusMsg.innerText = 'سەرکەوتوو بوو! دەچێتە ژوورەوە...';
            statusMsg.style.color = '#10b981';
            
            localStorage.setItem('sticker_auth', JSON.stringify({ username: 'admin', role: 'admin' }));
            
            setTimeout(() => {
                window.location.href = 'dashboard.html';
            }, 1000);
        } else if (username === 'user' && password === '1234') {
            statusMsg.innerText = 'سەرکەوتوو بوو! دەچێتە ژوورەوە...';
            statusMsg.style.color = '#10b981';
            
            localStorage.setItem('sticker_auth', JSON.stringify({ username: 'user', role: 'worker' }));
            
            setTimeout(() => {
                window.location.href = 'dashboard.html';
            }, 1000);
        } else {
            statusMsg.innerText = 'ناوی بەکارهێنەر یان وشەی نهێنی هەڵەیە.';
            statusMsg.className = 'status-msg error';
            statusMsg.style.color = '#ef4444';
        }
    }, 1000);
});
