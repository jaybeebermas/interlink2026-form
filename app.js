// ==========================================
// CONFIGURATION
// ==========================================
// Replace the string below with your deployed Google Apps Script Web App URL.
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbz16uQbSj_OOW6-8UaFABuP3AooR8pWGbCQxJSq2xdw6caTSk_6NzatOumAILTbvZxa1A/exec";

document.addEventListener("DOMContentLoaded", () => {
  // ==========================================
  // DOM ELEMENTS
  // ==========================================
  const form = document.getElementById('registrationForm');
  const schoolSelect = document.getElementById('school');
  const yearLevelSelect = document.getElementById('yearLevel');
  const programCSPCSelect = document.getElementById('programCSPC');
  const programUNAIRSelect = document.getElementById('programUNAIR');
  const specifyProgramInput = document.getElementById('specifyProgram');
  const specifyYearLevelInput = document.getElementById('specifyYearLevel');

  const programWrapper = document.getElementById('programWrapper');
  const cspcProgramContainer = document.getElementById('cspcProgramContainer');
  const unairProgramContainer = document.getElementById('unairProgramContainer');
  const specifyProgramWrapper = document.getElementById('specifyProgramWrapper');
  const specifyYearLevelWrapper = document.getElementById('specifyYearLevelWrapper');

  const statusOverlay = document.getElementById('statusOverlay');
  const statusLoading = document.getElementById('statusLoading');
  const statusSuccess = document.getElementById('statusSuccess');
  const statusError = document.getElementById('statusError');
  const successMessageText = document.getElementById('successMessage');
  const errorMessageText = document.getElementById('errorMessage');

  // ==========================================
  // CONDITIONAL VISIBILITY LOGIC
  // ==========================================
  function updateFormLogic() {
    const selectedSchool = schoolSelect.value;
    const selectedYearLevel = yearLevelSelect.value;
    const selectedProgramUNAIR = programUNAIRSelect.value;

    // 1. School logic
    if (selectedSchool === "CSPC") {
      // Show CSPC Program selection, hide UNAIR
      programWrapper.classList.add('show');
      cspcProgramContainer.classList.remove('d-none');
      programCSPCSelect.required = true;

      unairProgramContainer.classList.add('d-none');
      programUNAIRSelect.required = false;
      programUNAIRSelect.value = ""; // Clear unused selection

      // CSPC never specifies program
      specifyProgramWrapper.classList.remove('show');
      specifyProgramInput.required = false;
      specifyProgramInput.value = "";
    }
    else if (selectedSchool === "UNAIR") {
      // Show UNAIR Program selection, hide CSPC
      programWrapper.classList.add('show');
      unairProgramContainer.classList.remove('d-none');
      programUNAIRSelect.required = true;

      cspcProgramContainer.classList.add('d-none');
      programCSPCSelect.required = false;
      programCSPCSelect.value = ""; // Clear unused selection

      // Check if UNAIR Program is "Other" to show specifyProgram
      if (selectedProgramUNAIR === "Other") {
        specifyProgramWrapper.classList.add('show');
        specifyProgramInput.required = true;
      } else {
        specifyProgramWrapper.classList.remove('show');
        specifyProgramInput.required = false;
        specifyProgramInput.value = "";
      }
    }
    else {
      // No school selected - hide program elements
      programWrapper.classList.remove('show');
      cspcProgramContainer.classList.add('d-none');
      unairProgramContainer.classList.add('d-none');
      programCSPCSelect.required = false;
      programUNAIRSelect.required = false;
      programCSPCSelect.value = "";
      programUNAIRSelect.value = "";

      specifyProgramWrapper.classList.remove('show');
      specifyProgramInput.required = false;
      specifyProgramInput.value = "";
    }

    // 2. Year level specify logic
    if (selectedYearLevel === "Other") {
      specifyYearLevelWrapper.classList.add('show');
      specifyYearLevelInput.required = true;
    } else {
      specifyYearLevelWrapper.classList.remove('show');
      specifyYearLevelInput.required = false;
      specifyYearLevelInput.value = "";
    }
  }

  // Attach visibility triggers
  schoolSelect.addEventListener('change', updateFormLogic);
  programUNAIRSelect.addEventListener('change', updateFormLogic);
  yearLevelSelect.addEventListener('change', updateFormLogic);

  // Run once on load to ensure initial state is consistent
  updateFormLogic();

  // ==========================================
  // FORM SUBMISSION & REST API INTEGRATION
  // ==========================================
  form.addEventListener('submit', function (event) {
    event.preventDefault();

    // Enforce browser validation
    if (!form.checkValidity()) {
      event.stopPropagation();
      form.classList.add('was-validated');
      return;
    }

    // Show Loading Overlay
    statusOverlay.classList.remove('d-none');
    statusLoading.classList.remove('d-none');
    statusSuccess.classList.add('d-none');
    statusError.classList.add('d-none');

    // Validation Guard: Alert the user if the script URL hasn't been set
    if (WEB_APP_URL === "YOUR_SCRIPT_URL_HERE" || !WEB_APP_URL.startsWith("http")) {
      setTimeout(() => {
        statusLoading.classList.add('d-none');
        statusError.classList.remove('d-none');
        errorMessageText.textContent = "Configuration Error: Please update the WEB_APP_URL constant in app.js (line 5) with your actual Google Apps Script Web App URL.";
      }, 400);
      return;
    }

    // Compile data payload matching the backend schema.
    // Fields that are hidden / unused will submit empty strings.
    const payload = {
      firstName: form.firstName.value.trim(),
      middleInitial: form.middleInitial.value.trim().toUpperCase(),
      lastName: form.lastName.value.trim(),
      suffix: form.suffix.value.trim(),
      email: form.email.value.trim(),
      school: schoolSelect.value,
      programCSPC: programCSPCSelect.value,
      programUNAIR: programUNAIRSelect.value,
      specifyProgram: specifyProgramInput.value.trim(),
      yearLevel: yearLevelSelect.value,
      specifyYearLevel: specifyYearLevelInput.value.trim()
    };

    // We use text/plain for the POST requests to Google Apps Script. 
    // This bypasses browser preflight (OPTIONS) requests, avoiding redirection-related CORS blocks.
    fetch(WEB_APP_URL, {
      method: 'POST',
      mode: 'cors',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    })
      .then(response => {
        if (!response.ok) {
          throw new Error(`HTTP Error Status: ${response.status}`);
        }
        return response.json();
      })
      .then(data => {
        if (data.success) {
          // Success Response
          statusLoading.classList.add('d-none');
          statusSuccess.classList.remove('d-none');
          if (data.message) {
            successMessageText.textContent = data.message;
          }

          // Clear inputs & remove validation colors
          form.reset();
          form.classList.remove('was-validated');
          updateFormLogic(); // reset visibility to default
        } else {
          // Script reported failure (e.g. duplicate email)
          statusLoading.classList.add('d-none');
          statusError.classList.remove('d-none');
          errorMessageText.textContent = data.message || "Registration failed. Try resubmitting.";
        }
      })
      .catch(error => {
        console.error("Transmission Failure:", error);
        statusLoading.classList.add('d-none');
        statusError.classList.remove('d-none');
        errorMessageText.textContent = "Failed to communicate with database. Check your internet connection or URL settings.";
      });
  });

  // Expose status overlay functions globally to allow onclick handlers
  window.hideOverlay = function (resetRequired = false) {
    statusOverlay.classList.add('d-none');
    if (resetRequired) {
      form.classList.remove('was-validated');
    }
  };
});
