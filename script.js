"use strict"
document.addEventListener('DOMContentLoaded', function() {
    // Aspetta che il DOM sia completamente caricato

    // Aggiunge un listener all'input del file
    document.getElementById('excelFileInput').addEventListener('change', function(event) {
        let file = event.target.files[0]; // Prende il file selezionato dall'utente
        let reader = new FileReader(); // Crea un nuovo FileReader per leggere il contenuto del file
        
        reader.onload = function(e) {
            // Questa funzione viene chiamata quando il FileReader ha terminato di leggere il file
            let data = new Uint8Array(e.target.result); // Converti i dati del file in un array di byte
            let workbook = XLSX.read(data, {type: 'array'}); // Usa la libreria SheetJS (XLSX) per leggere l'array di byte come un workbook Excel
            let firstSheetName = workbook.SheetNames[0]; // Ottiene il nome del primo foglio del workbook
            let worksheet = workbook.Sheets[firstSheetName]; // Ottiene il primo foglio del workbook
            let html = convertSheetToHtml(worksheet); // Converte il foglio in una stringa HTML
            
            document.getElementById('output').textContent = html; // Mostra la stringa HTML come testo nell'elemento di output
        };
    
        reader.readAsArrayBuffer(file); // Inizia a leggere il contenuto del file
    });
    
    function convertSheetToHtml(worksheet) {
        // Funzione per convertire il foglio Excel in una tabella HTML
        let html = "<table>"; // Inizia con il tag di apertura della tabella
        let range = XLSX.utils.decode_range(worksheet['!ref']); // Decodifica il range di celle del foglio
    
        for(let R = range.s.r; R <= range.e.r; ++R) {
            // Itera su ogni riga del foglio
            html += "<tr>"; // Aggiunge il tag di apertura di una riga
            for(let C = range.s.c; C <= Math.min(range.e.c, 1); ++C) {
                // Itera sulle prime due colonne di ogni riga
                let cell_address = {c:C, r:R}; // Crea l'indirizzo della cella
                let cell_ref = XLSX.utils.encode_cell(cell_address); // Codifica l'indirizzo della cella in formato leggibile
                let cell = worksheet[cell_ref]; // Ottiene la cella dal foglio
                let cellValue = cell ? formatCellValue(cell.v) : ""; // Ottiene il valore della cella, applicando la formattazione se necessario
                html += "<td>" + cellValue + "</td>"; // Aggiunge il valore della cella nel tag td
            }
            html += "</tr>"; // Chiude il tag della riga
        }
        html += "</table>"; // Chiude il tag della tabella
        return html; // Restituisce la stringa HTML della tabella
    }
    
    function formatCellValue(value) {
        // Funzione per formattare i valori delle celle
        // Controlla se il valore è un numero che inizia con '0.'
        if ((typeof value === 'string' || typeof value === 'number') && /^0\.\d+$/.test(value.toString())) {
            let num = parseFloat(value) * 100; // Converte il valore in un numero e moltiplica per 100
            return num.toFixed(2) + '%'; // Formatta il numero come percentuale con due cifre decimali
        }
        return value; // Restituisce il valore non modificato se non corrisponde alla condizione
    }
});