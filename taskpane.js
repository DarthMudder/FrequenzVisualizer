/* global Office, document, requestAnimationFrame, window */

let currentFreq = 369;
let animationTime = 0;
let canvas, ctx;

// This function is called when Word or the browser environment initializes the Office JS framework
Office.onReady((info) => {
    
    // 1. Initialize the visualizer immediately so it works in BOTH the web browser and Microsoft Word
    initVisualizer();
    
    // 2. Execute Word-specific configurations only when running inside the desktop application
    if (info.host === Office.HostType.Word) {
        
        // --- AUTO-OPEN CONFIGURATION ---
        // Tells Word to automatically open this Add-in when the document is opened
        if (Office.context.document.settings) {
            Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
            Office.context.document.settings.saveAsync((result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.log("Failed to enable auto-open: " + result.error.message);
                } else {
                    console.log("Success! The Add-in will now open automatically with this document.");
                }
            });
        }
        // -------------------------------
    }
});

function initVisualizer() {
    const slider = document.getElementById("freqSlider");
    const valDisplay = document.getElementById("freqValue");
    canvas = document.getElementById("waveCanvas");
    ctx = canvas.getContext("2d");

    // Event Listener for the slider input
    slider.addEventListener("input", function() {
        currentFreq = parseInt(this.value, 10);
        valDisplay.textContent = currentFreq;
        updateDescription(currentFreq);
    });

    // Set the initial description based on the default value
    updateDescription(currentFreq);

    // Start the animation loop
    drawWave();
}

// Helper function for the HTML quick-select buttons
window.setFreq = function(val) {
    currentFreq = val;
    document.getElementById("freqSlider").value = val;
    document.getElementById("freqValue").textContent = val;
    updateDescription(val);
};

function updateDescription(freq) {
    const desc = document.getElementById("freqDescription");
    if (freq < 20) desc.innerHTML = `<strong>${freq} Hz (Infrasound / Brainwaves):</strong> Inaudible. Use Binaural Beats for effects like deep relaxation (e.g., Alpha waves at 10Hz).`;
    else if (freq === 369) desc.innerHTML = `<strong>369 Hz (Tesla / Balance):</strong> Represents a perfect balance between body and mind.`;
    else if (freq === 432) desc.innerHTML = `<strong>432 Hz (Earth Tone):</strong> Considered especially harmonious and can slightly lower the heart rate.`;
    else if (freq === 528) desc.innerHTML = `<strong>528 Hz (Healing / Love):</strong> Promotes deep inner peace and reduces stress.`;
    else if (freq >= 20 && freq < 250) desc.innerHTML = `<strong>${freq} Hz (Bass Range):</strong> Deep frequencies, often physically felt.`;
    else if (freq > 6000) desc.innerHTML = `<strong>${freq} Hz (Brilliance / High Pitch):</strong> Very high, shrill tones.`;
    else desc.innerHTML = `<strong>${freq} Hz (Mid-range):</strong> Normal hearing range, similar to human speech or instruments.`;
}

function drawWave() {
    // Clear the canvas for the next frame
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    
    // Draw the background
    ctx.fillStyle = '#201f1e';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // Draw the center line
    ctx.beginPath();
    ctx.moveTo(0, canvas.height / 2);
    ctx.lineTo(canvas.width, canvas.height / 2);
    ctx.strokeStyle = '#403e3d';
    ctx.lineWidth = 1;
    ctx.stroke();

    // Calculate and draw the sine wave
    ctx.beginPath();
    ctx.strokeStyle = '#0078d4'; // Microsoft Blue
    ctx.lineWidth = 3;

    // Adjust the visual density of the waves based on the frequency.
    let visualFrequency = Math.log10(currentFreq + 1) * 2; 

    for (let x = 0; x < canvas.width; x++) {
        // y = Amplitude * sin(Frequency * x + TimeShift)
        let y = (canvas.height / 2) + Math.sin((x * visualFrequency * 0.05) + animationTime) * (canvas.height / 3);
        
        if (x === 0) {
            ctx.moveTo(x, y);
        } else {
            ctx.lineTo(x, y);
        }
    }
    ctx.stroke();

    // Increment time for the next frame (faster animation at higher frequencies)
    animationTime += 0.05 + (currentFreq / 20000);

    // Request the next animation frame to create an endless loop
    requestAnimationFrame(drawWave);
}
