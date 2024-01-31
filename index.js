const { google } = require('googleapis');
const sheets = google.sheets('v4');

const spreadsheetId = '1D0rk5Je9Y8CWumqlU5w8j-FctJ8yCFvHPISW4OLRUts';
const range = 'A1:H';
const numberOfClasses = 60;

let situacion = '';
let passingGrade = 0;

const auth = new google.auth.GoogleAuth({
    keyFile: './credentials.json',
    scopes: [
        'https://www.googleapis.com/auth/spreadsheets.readonly',
        'https://www.googleapis.com/auth/spreadsheets'
    ],
});

async function getSpreadSheetValues() {
    try {
        const client = await auth.getClient();
        const response = await sheets.spreadsheets.values.get({
            auth: client,
            spreadsheetId,
            range,
        });

        const values = response.data.values;
        const headers = values[2];
        const headerIndex = {
            name: headers.indexOf('Aluno'),
            absences: headers.indexOf('Faltas'),
            p1: headers.indexOf('P1'),
            p2: headers.indexOf('P2'),
            p3: headers.indexOf('P3'),
            situacion: headers.indexOf('Situação'),
            pG: headers.indexOf('Nota para Aprovação Final')
        };

        const updateData = [];

        for(let i = 3; i < values.length; i++){
            const row = values[i];

            const name = row[headerIndex.name];
            const absences = parseInt(row[headerIndex.absences], 10);
            const p1 = parseFloat(row[headerIndex.p1]);
            const p2 = parseFloat(row[headerIndex.p2]);
            const p3 = parseFloat(row[headerIndex.p3]);
            
            calculateStudentsAverage(numberOfClasses, absences, p1, p2, p3);
            updateData.push([situacion, passingGrade]);
            console.log(`Name of the Student: ${name} \nNumber of Absences: ${absences} \nP1: ${p1}, P2: ${p2}, P3: ${p3} \nSituacion of the Student: ${situacion} \nPassing Grade: ${passingGrade}
            -------------------------------------------------------------------------`);
            await updateSpreadSheetValues(updateData);
        };
    }catch (error) {
        console.error('Error getting data from spreadsheet:',error.message);
    };
};

getSpreadSheetValues();

function calculateStudentsAverage(numberOfClasses, absences, p1, p2, p3){

    const average = (p1 + p2 + p3)/3;
    const percentualOfAbsences = (absences/numberOfClasses) * 100;

    if (percentualOfAbsences > 25) {
        situacion = 'Failed due to absence';
        passingGrade = 0;
    } else if (average < 50) {
        situacion = 'Failed by Grade';
        passingGrade = 0;
    } else if (average >= 50 && average < 70) {
        situacion = 'Final Exam';
        passingGrade = Math.round(2 * (70 - average));
        
    } else{
        situacion = 'Approved';
        passingGrade = 0;
    }

    return situacion, passingGrade.toFixed(2);
};

async function updateSpreadSheetValues(data){
    const updateRange = 'G4:H';
    const updateRangeWithRows = `${updateRange}${data.length + 3}`;

    const body = {
        values: data
    };

    try {
        const client = await auth.getClient();
        await sheets.spreadsheets.values.update({
            auth: client,
            spreadsheetId,
            range: updateRangeWithRows,
            valueInputOption: 'RAW',
            resource: body
        });
    }catch(err){
        console.log('Error when updating spreadsheet data:',err)
    }
};