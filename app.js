// Global variables and functions
let currentSheet = 'voorblad';

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    updateCurrentDate();
    showSheet('voorblad');
});

// Navigation function
function showSheet(sheetName) {
    // Hide all sheets
    const sheets = document.querySelectorAll('.sheet-container');
    sheets.forEach(sheet => sheet.classList.remove('active'));
    
    // Remove active class from all nav links
    const navLinks = document.querySelectorAll('.nav-link');
    navLinks.forEach(link => link.classList.remove('active'));
    
    // Show selected sheet
    const targetSheet = document.getElementById(sheetName);
    if (targetSheet) {
        targetSheet.classList.add('active');
        currentSheet = sheetName;
    }
    
    // Add active class to clicked nav link
    event.target.classList.add('active');
}

// Update current date in Afrikaans format
function updateCurrentDate() {
    const now = new Date();
    const afrikaansDate = formatDateAfrikaans(now);
    const dateDisplay = document.getElementById('current-date-display');
    if (dateDisplay) {
        dateDisplay.innerHTML = `<strong>Vandag se Datum:</strong> ${afrikaansDate}`;
    }
}

// Format date in Afrikaans
function formatDateAfrikaans(date) {
    const months = [
        'Januarie', 'Februarie', 'Maart', 'April', 'Mei', 'Junie',
        'Julie', 'Augustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    
    const day = date.getDate();
    const month = months[date.getMonth()];
    const year = date.getFullYear();
    
    return `${day} ${month} ${year}`;
}

// Placeholder functions for all macros (to be implemented in subsequent steps)
function showform() {
    alert('showform() - To be implemented');
}

function SortFamilyMembers() {
    alert('SortFamilyMembers() - To be implemented');
}

function CreateRegisterPDF() {
    alert('CreateRegisterPDF() - To be implemented');
}

function PrintToPDF_Landscape() {
    alert('PrintToPDF_Landscape() - To be implemented');
}

function DeleteRowsBelowseven() {
    alert('DeleteRowsBelowseven() - To be implemented');
}

function PrintToPDF_Landscape2() {
    alert('PrintToPDF_Landscape2() - To be implemented');
}

function DeleteRowsBelowTen() {
    alert('DeleteRowsBelowTen() - To be implemented');
}

function PrintToPDF_Landscape4() {
    alert('PrintToPDF_Landscape4() - To be implemented');
}

function CopyDate() {
    alert('CopyDate() - To be implemented');
}

function CopyDateVerjaarsade() {
    alert('CopyDateVerjaarsade() - To be implemented');
}

function CopyDateAndDisplayMessageBox() {
    alert('CopyDateAndDisplayMessageBox() - To be implemented');
}

function CopyAnniversary() {
    alert('CopyAnniversary() - To be implemented');
}

function SendContactEmails() {
    alert('SendContactEmails() - To be implemented');
}

function UpdateEPosSheet() {
    alert('UpdateEPosSheet() - To be implemented');
}
