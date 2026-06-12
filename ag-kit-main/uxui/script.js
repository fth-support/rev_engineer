(function () {
    const toggle = document.getElementById('theme-toggle');
    if (!toggle) return;

    function setTheme(dark) {
        document.body.classList.toggle('dark', dark);
        localStorage.setItem('bento-theme', dark ? 'dark' : 'light');
    }

    toggle.addEventListener('click', () => {
        setTheme(!document.body.classList.contains('dark'));
    });

    const saved = localStorage.getItem('bento-theme');
    if (saved === 'dark') setTheme(true);
})();
