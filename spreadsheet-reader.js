const axios = require('axios'); // library to work with HTTP Requests
const { google } = require('googleapis'); // google api library

//------------------------- Authentication and obtaining spreadsheet credentials -----------------------------

// Spreedsheet URL
const sheetUrl = 'https://docs.google.com/spreadsheets/d/1bJqNW7MaVojNTA-F2hY7_1ETiatPv0deXp1PjSkmoLQ/edit?usp=sharing';


// Extracting the ID from the spreadsheet URL
const sheetId = getSheetId(sheetUrl);


// Authenticating Google service account credentials to modify the spreadsheet cell values
const auth = new google.auth.GoogleAuth({
    keyFile: './credentials.json', // JSON file with the service account credentials {key, ID, and email}
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});


// Method to extract the ID -> splitting the URL at "/", only the 5th element of the array is of interest (it's the ID code)
function getSheetId(sheetUrl) {
    const parts = sheetUrl.split('/');
    const id = parts[5];
    return id;
}

// Method to get the cell values from the spreadsheet, in JSON format for better visualization and organization
function getInfo(infoJSON) {
    // Formula to organize spreadsheet values from the request and arrange them in arrays
    const info = JSON.parse(infoJSON.substr(infoJSON.indexOf('(') + 1, infoJSON.lastIndexOf(')') - infoJSON.indexOf('(') - 1));
    
    /* Identifies the table rows and maps each cell to get it's value, if the cell is empty 
    it's assigned the value "null" to the cell position in the future data array */
    const rows = info.table.rows;
    const values = rows.map(row => {
        return row.c.map(cell => {
            return cell ? cell.v : null;
        });
    });
    return values;
}

//-------------------------------------------------------------------------------------------------------------------------


/* Through axios, an HTTP request is made to return the values ​​contained in the spreadsheet.
 
    .then -> the request returns a JSON which is converted into an array according to the rows by the function
    getInfo() and those arrays are inserted into "students_array", a way to concentrate and organize the data of all the 
    students in just one array

    .catch -> handles any connection error from the request
 
 */

axios.get(`https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json`)
    .then(response => {
        const data = response.data;
        const students_array = getInfo(data);
        status(average(students_array), absences(students_array), students_array);

    })
    .catch(error => {
        console.error('Error:', error.response.data);
    });



// ---------------------------------- Spreadsheet manipulation functions ------------------------------


/* Method for calculating student averages 

    "studentsAverages" -> array that will contain the final averages of all students.
    Through "for" loop, the array containing the data of all students is sectioned with the method
    slice, only the elements from the 3rd to the 6th position (grade values) are taken from this array, and through the
    reduce method they are condensed into one single variable, where the average value is calculated within the range (0 - 10).
    Finally, this value is inserted into the array of averages for easier organization.
    Subsequently, the array is returned to be used as a parameter in the function that establishes the student's Status

*/
function average(students) {
    let studentsAverages = [];
    for (let student of students) {
        const studentAverage = (((student.slice(3, 6).reduce((sum, grade) => sum + grade, 0)) / 3) / 10);
        studentsAverages.push(studentAverage.toFixed(2));
    }
    return studentsAverages;
}


/* Method for calculating the percentage of student absences 

    "absences" -> array that will contain the absences of all students.
    Through "for" loop, the array containing the data of all students is traveled through and the value from
    the 2nd position (student's absences quantity) is extracted and inserted into the absences array.
    Subsequently, the array is returned to be used as a parameter in the function that establishes the student's Status.

*/
function absences(students) {
    let absences = [];
    for (student of students) {
        absences.push(student[2]);
    }
    return absences;
}



/* Method for calculating the Final Approval Grade 
    
        Calculation formula ->  (average + naf)/2) >= 5)

        Through mathematical manipulation, we get: naf >= 10 - average
        The average value is passed as a parameter and is used for the grade calculation 
        and returned to be inserted into the spreadsheet cell.
*/
function approval(average) {
    let naf = 10 - average;
    return naf;

}




/* Method for updating the values ​​of spreadsheet cells 

    This function is a try-catch block, where in try the connection is established through the 
    google auth client with the spreadsheet and allow the insertion of data into the cells.
    This authentication asks for the spreadsheet ID which has already been extracted earlier by the getSheetId() function,
    the cell to be modified, the input option, and the value to be applied.

*/
async function updateCell(spreadsheetId, cell, value) {
    try {
        const client = await auth.getClient();
        const sheets = google.sheets({ version: 'v4', auth: client });
        const response = await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: cell,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[value]]
            }
        });
        console.log('Sucess:', response.data);
    } catch (error) {
        console.error('Error:', error);
    }
}


/* Method for managing the student's Status 

    This function receives as a parameter the return of the average and absence functions, as well as the array of students itself 
    to perform calculations.

    The Status column is traveled throught and it's first analyzed whether the student has enough attendance 
    to be approved, through If - Else, if the student does not have enough attendance they are already classified as Failure due to attendance.

    If the student meets the attendance requirements, the grade is analyzed and classified as (Final Exam,
        Failure due to grade or Approved).

    And according to each category, the corresponding cell is updated, and the corresponding cell number of the next column
    if the student has already been (Approved, Failed by Grade, or Failed by Absence) receives the value 0. If the student's status is Final Exam,
    the corresponding cell of the next column will be updated with the necessary value for the Final Approval Grade.


*/
async function status(average, absences, students) {
    const absence_limit = 15 // total semester classes = 60 -> 25% of 60 classes -> 15 classes is the limit of absences for students
    let pos = 0;
    let rowNumber = 4;

    while (pos < students.length) {
        let cell = 'G' + rowNumber;
        let nextCell = 'H' + rowNumber;

        if (absences[pos] <= absence_limit) {

            if (average[pos] < 5) {
                await updateCell(sheetId, cell, "Reprovado por Nota");
                await updateCell(sheetId, nextCell, "0");

            } else if (5 <= average[pos] && average[pos] < 7) {
                await updateCell(sheetId, cell, "Exame Final");
                await updateCell(sheetId, nextCell, approval(average[pos]));

            } else {
                await updateCell(sheetId, cell, "Aprovado");
                await updateCell(sheetId, nextCell, "0");
            }

        } else {
            await updateCell(sheetId, cell, "Reprovado por Falta");
            await updateCell(sheetId, nextCell, "0");
        }
        rowNumber++;
        pos++;
    }
}
