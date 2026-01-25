document.addEventListener('DOMContentLoaded', function() {
    const body = document.body;
    const greetingEl = document.getElementById('greeting');
    const themeToggle = document.getElementById('theme-toggle');
    const allThemes = ['theme-morning', 'theme-afternoon', 'theme-evening', 'theme-dark-mode'];

    const btnMorning = document.getElementById('btn-morning-theme');
    const btnAfternoon = document.getElementById('btn-afternoon-theme');
    const btnEvening = document.getElementById('btn-evening-theme');

    // Centralized configuration for all themes.
    // The array is ordered by start time in descending order.
    const THEME_DEFINITIONS = [
        { name: 'night', startHour: 21, greeting: 'Hello Night Owl! 🦉', cssClass: 'theme-dark-mode' },
        { name: 'evening', startHour: 17, greeting: 'Good Evening! 🌆', cssClass: 'theme-evening' },
        { name: 'afternoon', startHour: 12, greeting: 'Good Afternoon! ☀️', cssClass: 'theme-afternoon' },
        { name: 'morning', startHour: 6, greeting: 'Good Morning! 🌅', cssClass: 'theme-morning' }
    ];
    
    // Fallback theme for hours before the first defined start time (e.g., before 6 AM).
    const FALLBACK_THEME = THEME_DEFINITIONS[0]; // Night

    /**
     * Gets the complete theme configuration object based on the current hour.
     * This is the single source of truth for time-based theme logic.
     * @returns {object} The configuration object for the current theme.
     */
    function getCurrentThemeConfig() {
        const currentHour = new Date().getHours();
        // Find the first theme where the current hour is greater than or equal to its start time.
        const theme = THEME_DEFINITIONS.find(t => currentHour >= t.startHour);
        return theme || FALLBACK_THEME;
    }

    let isInitialLoad = true;

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

    /**
     * Updates the greeting text based on the provided theme CSS class, with a flip animation.
     * @param {string} themeClassName The CSS class of the currently applied theme.
     */
    function updateGreeting(themeClassName) {
        if (!greetingEl) return;

        const themeConfig = THEME_DEFINITIONS.find(t => t.cssClass === themeClassName);
        const newGreetingText = themeConfig ? themeConfig.greeting : getCurrentThemeConfig().greeting;

        if (isInitialLoad) {
            greetingEl.textContent = newGreetingText;
            isInitialLoad = false;
            return;
        }

        // Prevent animation if text is not changing
        if (greetingEl.textContent === newGreetingText) {
            return;
        }

        greetingEl.classList.add('flipping-out');

        setTimeout(() => {
            greetingEl.textContent = newGreetingText;
            greetingEl.classList.remove('flipping-out');
            greetingEl.classList.add('flipping-in');

            // Clean up the class after animation ends
            greetingEl.addEventListener('animationend', () => {
                greetingEl.classList.remove('flipping-in');
            }, { once: true });

        }, 300); // Must match the first animation's duration
    }

    /**
     * Determines the appropriate theme CSS class based on the time of day.
     * @returns {string} The CSS class for the time-based theme.
     */
    function getTimeBasedTheme() {
        return getCurrentThemeConfig().cssClass;
    }

    /**
     * Gets the appropriate "light" theme (morning, afternoon, or evening) based on the time.
     * This is used when toggling dark mode off.
     * @returns {string} The CSS class for the time-based light theme.
     */
    function getLightTimeBasedTheme() {
        const config = getCurrentThemeConfig();
        // If the current time-based theme is dark mode, default to the 'evening' theme.
        if (config.name === 'night') {
            return THEME_DEFINITIONS.find(t => t.name === 'evening').cssClass;
        }
        return config.cssClass;
    }

    if (btnMorning) {
        btnMorning.addEventListener('click', () => {
            applyTheme(THEME_DEFINITIONS.find(t => t.name === 'morning').cssClass);
        });
    }
    if (btnAfternoon) {
        btnAfternoon.addEventListener('click', () => {
            applyTheme(THEME_DEFINITIONS.find(t => t.name === 'afternoon').cssClass);
        });
    }
    if (btnEvening) {
        btnEvening.addEventListener('click', () => {
            applyTheme(THEME_DEFINITIONS.find(t => t.name === 'evening').cssClass);
        });
    }

    if (themeToggle) {
        themeToggle.addEventListener('change', function() {
            if (this.checked) {
                applyTheme(THEME_DEFINITIONS.find(t => t.name === 'night').cssClass);
            } else {
                applyTheme(getLightTimeBasedTheme());
            }
        });
    }

    function initializeTheme() {
        applyTheme(getTimeBasedTheme());
    }

    initializeTheme();
});
