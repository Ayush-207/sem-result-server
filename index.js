import express from 'express';
// import fs from 'fs-extra';
import excelToJson from 'convert-excel-to-json';
import cors from 'cors';
import helmet from 'helmet';

const app = express();
app.use(helmet());
app.use(helmet.crossOriginResourcePolicy({ policy: "cross-origin" }));
app.use(cors());
const PORT = 3001;

// const str = ["Sub Cr","GR","GP","CRP"];
// const keys = {};
// for(let i=7; i<)

var filepath = 'assets/result5sem.xlsx';
let key = {
    A: 'Sl',
    B: 'Roll No',
    C: "Name\nFather's Name",
    D: 'Sub Cr',
    E: 'GR',
    F: 'GP',
    G: 'CRP',
    H: 'Sub Cr',
    I: 'GR',
    J: 'GP',
    K: 'CRP',
    L: 'Sub Cr',
    M: 'GR',
    N: 'GP',
    O: 'CRP',
    P: 'Sub Cr',
    Q: 'GR',
    R: 'GP',
    S: 'CRP',
    T: 'Sub Cr',
    U: 'GR',
    V: 'GP',
    W: 'CRP',
    X: 'Sub Cr',
    Y: 'GR',
    Z: 'GP',
    AA: 'CRP',
    AB: 'Sub Cr',
    AC: 'GR',
    AD: 'GP',
    AE: 'CRP',
    AF: 'TOT CR',
    AG: 'TOT CRP',
    AH: 'SGPA',
    AI: 'CS'
};
const data = excelToJson({
    sourceFile: filepath,
    header: { rows: 2 },
});

function getCourseStats(coursecode, branchcode = '') {
    // console.log([coursecode, branchcode]);
    let totalstudents = 0;
    let result = {};
    for (let i = 1; i < 38; i++) {
        const tablename = 'Table ' + i;
        const tabledata = data[tablename];
        for (let j = 0; j < tabledata.length; j++) {
            let values = Object.values(tabledata[j]);

            if (branchcode == '' || tabledata[j].B && tabledata[j].B.includes(branchcode.toUpperCase()) || tabledata[j].C && tabledata[j].C.includes(branchcode.toUpperCase())) {
                if (values.indexOf(coursecode.toUpperCase() + " 4") > -1) {
                    for (let k = values.indexOf(coursecode.toUpperCase() + " 4"); k < values.length; k++) {
                        if (typeof (values[k]) == 'number') {
                            if (result[values[k]])
                                result[values[k]]++;
                            else result[values[k]] = 1;
                            totalstudents++;
                            break;
                        }
                    }
                }

            }
        }
    }

    // console.log(result);
    // console.log(totalstudents);
    return [result, totalstudents];
}

function getOverallStats(branchcode = '') {
    let totalstudents = 0;
    let result = {};
    let keysk = Object.keys(key);
    for (let i = 1; i < 38; i++) {
        const tablename = 'Table ' + i;
        const tabledata = data[tablename];
        for (let j = 0; j < tabledata.length; j++) {
            let values = Object.values(tabledata[j]);
            if (tabledata[j].B && tabledata[j].B.includes("2021" + branchcode.toUpperCase()) || tabledata[j].C && tabledata[j].C.includes("2021" + branchcode.toUpperCase())) {
                if (typeof (values[values.length - 2]) == 'number') {
                    if (result[Math.floor(values[values.length - 2])]) result[Math.floor(values[values.length - 2])]++;
                    else result[Math.floor(values[values.length - 2])] = 1;
                    totalstudents++;
                    if (Math.floor(values[values.length - 2]) == 10) {
                        console.log(tabledata[j]);
                    }
                }
            }
        }
    }

    // console.log(result);
    // console.log(totalstudents);
    return [result, totalstudents];
}

// getOverallStats();

// getCourseStats('MEMEC14', 'UME');
// console.log(data['Table 28'][127]['S'].includes('EONS006'));

function getResult(rollno) {
    let flag = false;
    let result = {};
    for (let i = 1; i < 38; i++) {
        const tablename = "Table " + i;
        const tableData = data[tablename];
        for (let j = 0; j < tableData.length; j++) {
            if (tableData[j].B == rollno || tableData[j].C == rollno || tableData[j].D == rollno) {
                flag = true;
                result = tableData[j];
                break;
            }
        }
        if (flag) break;
    }

    let keysk = Object.values(key);
    let values = Object.values(result);

    return [keysk, values];
}

// console.log(getResult(rolln));

app.get('/getresult', (req, res) => {
    const rolln = req.query.rollno;
    const result = getResult(rolln);
    res.status(200).send(result);
});

app.get('/getcoursestats', (req, res) => {
    const coursecode = req.query.coursecode;
    const branchcode = req.query.branchcode;
    const stats = getCourseStats(coursecode, branchcode);

    res.status(200).send(stats);
});

app.get('/getoverallstats', (req, res) => {
    const branchcode = req.query.branchcode;
    const stats = getOverallStats(branchcode);
    res.status(200).send(stats);
});

app.listen(PORT, (error) => {
    if (!error) {
        console.log("Server is running on " + PORT);
    }
    else {
        console.log("Error occurred, server can't start", error);
    }
})