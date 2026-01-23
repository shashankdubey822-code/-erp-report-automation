var sleepTimer; // This is now used as mouseIdleTimeout in the new logic
var angryTimeout;
var sleepModeOn = false;
var mouseX = 0; // Global variable for mouse X position
var mouseY = 0; // Global variable for mouse Y position

// Get the owl container once to avoid repeated DOM queries
var owlContainer = document.querySelector('.owl-container');

// Functions to retrieve DOM elements
function getPupilLeft() { return document.getElementById("pupilLeft"); }
function getPupilRight() { return document.getElementById("pupilRight"); }
function getEyeLeft() { return document.getElementById("eyeLeft"); }
function getEyeRight() { return document.getElementById("eyeRight"); }
function getUpperLids() { return document.getElementsByClassName('eye-lid-up'); }
function getDownLids() { return document.getElementsByClassName('eye-lid-down'); }
function getWingLeft() { return document.getElementsByClassName('wing-left')[0]; }
function getWingRight() { return document.getElementsByClassName('wing-right')[0]; }


// Timer for mouse inactivity
var mouseIdleTimeout;

// Pupil Max Movement - This defines the range of the pupil's linear movement
var pupilMaxMove = 25; 

// Base offset for pupil to be centered when no movement
var pupilBaseOffset = 35; // Pupil is 100x100 in 170x170 eye, so 35px from edge centers it.


function calculatePupilOffset(eyeCenterX, eyeCenterY, pupilElement) {
    var dx_mouse_to_eye_center = mouseX - eyeCenterX;
    var dy_mouse_to_eye_center = mouseY - eyeCenterY;
    
    // Limit pupil movement to pupilMaxMove along X and Y axes directly
    var pupilOffsetX = Math.max(-pupilMaxMove, Math.min(pupilMaxMove, dx_mouse_to_eye_center));
    var pupilOffsetY = Math.max(-pupilMaxMove, Math.min(pupilMaxMove, dy_mouse_to_eye_center));

    // Apply offset from the base centered position
    pupilElement.style.left = (pupilBaseOffset + pupilOffsetX) + "px";
    pupilElement.style.top = (pupilBaseOffset + pupilOffsetY) + "px";
}

function mouseMove(e) {
    // console.log('mouse moved');
    clearTimeout(mouseIdleTimeout); // Clear any existing idle timer
    
    // If in sleep mode, wake up the owl
    if (sleepModeOn) {
        setDefaultAnimation();
        sleepModeOn = false;
    }

    // Reset idle timer
    mouseIdleTimeout = setTimeout(startIdleSleep, 1000); // 1 second idle for sleep

    var pupilLeft = getPupilLeft();
    var pupilRight = getPupilRight();
    var eyeLeft = getEyeLeft();
    var eyeRight = getEyeRight();

    if (!pupilLeft || !pupilRight || !eyeLeft || !eyeRight) {
        // Elements not found, maybe owl is not displayed or not fully loaded
        return;
    }

    // Calculate the center of the eye elements relative to the viewport (these are the static white eyes)
    var leftEyeRect = eyeLeft.getBoundingClientRect();
    var rightEyeRect = eyeRight.getBoundingClientRect();

    var leftEyeCenterX = leftEyeRect.left + (leftEyeRect.width / 2);
    var leftEyeCenterY = leftEyeRect.top + (leftEyeRect.height / 2);

    var rightEyeCenterX = rightEyeRect.left + (rightEyeRect.width / 2);
    var rightEyeCenterY = rightEyeRect.top + (rightEyeRect.height / 2);

    // Mouse position (update global variables)
    mouseX = e.clientX;
    mouseY = e.clientY;

    // --- White Eye (eyeLeft, eyeRight) remain STATIC ---
    // Clear any transforms that might cause movement or rotation of the white eye part
    eyeLeft.style.transform = ''; 
    eyeRight.style.transform = ''; 

    // --- Pupil Movement (Inner Eye) ---
    calculatePupilOffset(leftEyeCenterX, leftEyeCenterY, pupilLeft);
    calculatePupilOffset(rightEyeCenterX, rightEyeCenterY, pupilRight);
}

function startIdleSleep() {
    // console.log('Mouse idle, initiating sleep');
    if (!sleepModeOn) {
        sleepModeOn = true;
        setAngryOrSleepAnimation('sleep'); // Explicitly call sleep animation
    }
}


function mouseOut(e) { // 'el' parameter is no longer needed
    // console.log('mouse left window');
    // Start idle timer when mouse leaves the document
    clearTimeout(mouseIdleTimeout);
    mouseIdleTimeout = setTimeout(startIdleSleep, 100); // Short delay to prevent accidental sleep on quick re-entry
}


function setAngryOrSleepAnimation(mode) {
    var upperLid = getUpperLids();
    for (var i = 0; i < upperLid.length; i++) {
        upperLid[i].style.animation = mode === 'angry' ? 'lid-up-angry 0.5s 1 forwards' : 'lid-up-sleep 2s 1 forwards';
    }
    var downLid = getDownLids();
    for (var i = 0; i < downLid.length; i++) {
        downLid[i].style.animation = mode === 'angry' ? 'lid-down-angry 0.5s 1 forwards' : 'lid-down-sleep 2s 1 forwards';
    }
    var wingLeft = getWingLeft();
    if (wingLeft) {
        wingLeft.style.transition = 'transform 1s';
        wingLeft.style.animation = 'left-wing-sleep 1s 1 forwards';
    }
    var wingRight = getWingRight();
    if (wingRight) {
        wingRight.style.transition = 'transform 1s';
        wingRight.style.animation = 'right-wing-sleep 1s 1 forwards';
    }
}

function setDefaultAnimation() {
    var upperLid = getUpperLids();
    for (var i = 0; i < upperLid.length; i++) {
        upperLid[i].style.animation = 'lid-up-blink 2.5s infinite';
    }
    var downLid = getDownLids();
    for (var i = 0; i < downLid.length; i++) {
        downLid[i].style.animation = 'lid-down-blink 2.5s infinite';
    }
    var wingLeft = getWingLeft();
    if (wingLeft) {
        wingLeft.style.animation = 'left-wing-move 5s 0.5s infinite';
        wingLeft.style.transition = '';
    }
    var wingRight = getWingRight();
    if (wingRight) {
        wingRight.style.transition = '';
        wingRight.style.animation = 'right-wing-move 5s 0.5s infinite';
    }
}

function angryStart(event) {
    event.stopPropagation();
    // console.log('angry triggered');
    clearTimeout(angryTimeout);
    setAngryOrSleepAnimation('angry');
    angryTimeout = setTimeout(function () {
        // console.log('angry timeout');
        setDefaultAnimation();
        // After angry mode, set back to idle timer
        clearTimeout(mouseIdleTimeout);
        mouseIdleTimeout = setTimeout(startIdleSleep, 1000);
    }, 3000);
}

// Function to dynamically adjust margin for owl spacing
function adjustOwlSpacing() {
    const nightJoySection = document.getElementById('night-mode-joyful-section');
    if (!nightJoySection || nightJoySection.style.display === 'none') {
        return;
    }

    const textAndOwlParent = nightJoySection.querySelector('.text-center.p-5.rounded');
    const h2Element = textAndOwlParent ? textAndOwlParent.querySelector('h2') : null;
    const pElement = textAndOwlParent ? textAndOwlParent.querySelector('p.lead') : null;
    const owlAnimationDiv = textAndOwlParent ? textAndOwlParent.querySelector('.night-joy-animation') : null;

    if (!h2Element || !pElement || !owlAnimationDiv) {
        return;
    }

    // Calculate the combined height of the text block (h2 and p elements), including their margins
    let combinedTextContentHeight = h2Element.offsetHeight;
    combinedTextContentHeight += parseInt(window.getComputedStyle(h2Element).marginTop) + parseInt(window.getComputedStyle(h2Element).marginBottom);
    
    combinedTextContentHeight += pElement.offsetHeight;
    combinedTextContentHeight += parseInt(window.getComputedStyle(pElement).marginTop) + parseInt(window.getComputedStyle(pElement).marginBottom);

    const minGapDesired = 30; // Minimum desired pixel gap between the text block and the owl

    // Get the current vertical position of the owlAnimationDiv relative to its parent (textAndOwlParent)
    // This value includes any current margin-top.
    const currentOwlOffsetTop = owlAnimationDiv.offsetTop;
    
    // Calculate the required top position for the owlAnimationDiv to ensure the desired gap
    // This is where its top edge (including its own margin-top) should be.
    const requiredOwlOffsetTop = combinedTextContentHeight + minGapDesired;

    // Calculate the difference. This difference is what needs to be effectively added to the
    // owlAnimationDiv's current margin-top to reach the required position.
    const currentMarginTop = parseInt(window.getComputedStyle(owlAnimationDiv).marginTop) || 0;

    // The new margin-top will be `currentMarginTop + (requiredOwlOffsetTop - currentOwlOffsetTop)`
    let newMarginTop = currentMarginTop + (requiredOwlOffsetTop - currentOwlOffsetTop);
    
    // Ensure newMarginTop is not negative and has a reasonable base value
    const baseMarginTopForOwlDiv = 20; // A base margin even if no extra gap is needed
    owlAnimationDiv.style.marginTop = `${Math.max(baseMarginTopForOwlDiv, newMarginTop)}px`;
}


// Add global event listeners when the DOM is fully loaded
document.addEventListener('DOMContentLoaded', function() {
    // Only attach global listeners if owlContainer is found
    if (owlContainer) {
        window.addEventListener('mousemove', mouseMove);
        document.body.addEventListener('mouseleave', mouseOut);
        // Initial setup for idle sleep
        mouseIdleTimeout = setTimeout(startIdleSleep, 1000);
    }
    
    // Call adjustOwlSpacing on initial load
    adjustOwlSpacing();
    // Also call on window resize
    window.addEventListener('resize', adjustOwlSpacing);

    // Ensure adjustOwlSpacing is called when theme changes to dark mode
    const themeToggle = document.getElementById('theme-toggle');
    if (themeToggle) {
        themeToggle.addEventListener('change', function() {
            if (this.checked) { // Dark mode activated
                adjustOwlSpacing();
            }
        });
    }

    // Also call when morning/afternoon/evening buttons are clicked
    const btnMorning = document.getElementById('btn-morning-theme');
    const btnAfternoon = document.getElementById('btn-afternoon-theme');
    const btnEvening = document.getElementById('btn-evening-theme');

    if (btnMorning) {
        btnMorning.addEventListener('click', adjustOwlSpacing);
    }
    if (btnAfternoon) {
        btnAfternoon.addEventListener('click', adjustOwlSpacing);
    }
    if (btnEvening) {
        btnEvening.addEventListener('click', adjustOwlSpacing);
    }
});