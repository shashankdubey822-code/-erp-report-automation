document.addEventListener('DOMContentLoaded', function() {
    const body = document.body;
    const greetingEl = document.getElementById('greeting');
    const themeToggle = document.getElementById('theme-toggle');
    const allThemes = ['theme-morning', 'theme-afternoon', 'theme-evening', 'theme-dark-mode'];

    const btnMorning = document.getElementById('btn-morning-theme');
    const btnAfternoon = document.getElementById('btn-afternoon-theme');
    const btnEvening = document.getElementById('btn-evening-theme');

    function applyTheme(themeClassName) {
        body.classList.remove(...allThemes);
        if (themeClassName) {
            body.classList.add(themeClassName);
        }
        updateGreeting(themeClassName);
        if (themeToggle) {
            themeToggle.checked = (themeClassName === 'theme-dark-mode');
        }
    }

    function updateGreeting(themeClassName) {
        if (!greetingEl) return;
        let greetingText = '';

        switch (themeClassName) {
            case 'theme-morning':
                greetingText = 'Good Morning! 🌅';
                break;
            case 'theme-afternoon':
                greetingText = 'Good Afternoon! ☀️';
                break;
            case 'theme-evening':
                greetingText = 'Good Evening! 🌆';
                break;
            case 'theme-dark-mode':
                greetingText = 'Hello Night Owl! 🦉';
                break;
            default:
                // Fallback in case themeClassName is null or unexpected
                const hour = new Date().getHours();
                if (hour >= 6 && hour < 13) {
                    greetingText = 'Good Morning! 🌅';
                } else if (hour >= 13 && hour < 17) {
                    greetingText = 'Good Afternoon! ☀️';
                } else if (hour >= 17 && hour < 21) {
                    greetingText = 'Good Evening! 🌆';
                } else {
                    greetingText = 'Hello Night Owl! 🦉';
                }
                break;
        }
        greetingEl.textContent = greetingText;
    }

    function getTimeBasedTheme() {
        const hour = new Date().getHours();
        if (hour >= 6 && hour < 13) {
            return 'theme-morning';
        } else if (hour >= 13 && hour < 17) {
            return 'theme-afternoon';
        } else if (hour >= 17 && hour < 21) {
            return 'theme-evening';
        } else {
            return 'theme-dark-mode';
        }
    }

    function getLightTimeBasedTheme() {
        const hour = new Date().getHours();
        if (hour >= 6 && hour < 13) {
            return 'theme-morning';
        } else if (hour >= 13 && hour < 17) {
            return 'theme-afternoon';
        } else {
            return 'theme-evening';
        }
    }

    if (btnMorning) {
        btnMorning.addEventListener('click', () => {
            localStorage.setItem('manualTheme', 'morning');
            applyTheme('theme-morning');
        });
    }
    if (btnAfternoon) {
        btnAfternoon.addEventListener('click', () => {
            localStorage.setItem('manualTheme', 'afternoon');
            applyTheme('theme-afternoon');
        });
    }
    if (btnEvening) {
        btnEvening.addEventListener('click', () => {
            localStorage.setItem('manualTheme', 'evening');
            applyTheme('theme-evening');
        });
    }

    if (themeToggle) {
        themeToggle.addEventListener('change', function() {
            if (this.checked) {
                localStorage.setItem('manualTheme', 'dark-mode');
                applyTheme('theme-dark-mode');
            } else {
                localStorage.removeItem('manualTheme');
                applyTheme(getLightTimeBasedTheme());
            }
        });
    }

    function initializeTheme() {
        const manualTheme = localStorage.getItem('manualTheme');

        if (manualTheme === 'dark-mode') {
            applyTheme('theme-dark-mode');
        } else if (manualTheme === 'morning') {
            applyTheme('theme-morning');
        } else if (manualTheme === 'afternoon') {
            applyTheme('theme-afternoon');
        } else if (manualTheme === 'evening') {
            applyTheme('theme-evening');
        } else {
            applyTheme(getTimeBasedTheme());
        }
    }

    initializeTheme();
});
