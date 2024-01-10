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

var filepath3 = 'assets/result5sem.xlsx';
var filepath1 = 'assets/result1sem.xlsx';
var filepath2 = 'assets/result3sem.xlsx';
var filepath4 = 'assets/result7sem.xlsx';

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
const data1 = excelToJson({
    sourceFile: filepath1,
    header: { rows: 2 },
});
const data2 = excelToJson({
    sourceFile: filepath2,
    header: { rows: 2 },
});
const data3 = excelToJson({
    sourceFile: filepath3,
    header: { rows: 2 },
});
const data4 = excelToJson({
    sourceFile: filepath4,
    header: { rows: 2 },
});

const data = [data1, data2, data3, data4];
const pages = [49, 49, 37, 36];

function getCourseStats(coursecode, branchcode = '', yoa) {
    // console.log([coursecode, branchcode]);
    if (coursecode == '' && yoa == '') {
        return [{}, 0];
    }
    branchcode = branchcode == '' ? '' : branchcode.toUpperCase();
    let totalstudents = 0;
    let result = {};
    for (let i = 1; i <= pages['2023' - yoa]; i++) {
        const tablename = 'Table ' + i;
        const tabledata = data['2023' - yoa][tablename];
        for (let j = 0; j < tabledata.length; j++) {
            let values = Object.values(tabledata[j]);

            if (branchcode == '' || tabledata[j].B && tabledata[j].B.includes(branchcode) || tabledata[j].C && tabledata[j].C.includes(branchcode)) {
                let index = values.length;
                for (let k = 0; k < values.length; k++) {
                    if (values[k].toString().includes(coursecode.toUpperCase())) {
                        index = k;
                    }
                }

                for (let k = index; k < values.length; k++) {
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

    // console.log(result);
    // console.log(totalstudents);
    return [result, totalstudents];
}

// console.log(getCourseStats('FENM012', 'UIT', '2022'));

function getOverallStats(branchcode = '', yoa) {
    if (yoa == '') {
        return [{}, 0];
    }
    branchcode = branchcode == '' ? '' : branchcode.toUpperCase();
    let totalstudents = 0;
    let result = {};
    let keysk = Object.keys(key);
    for (let i = 1; i <= pages['2023' - yoa]; i++) {
        const tablename = 'Table ' + i;
        const tabledata = data['2023' - yoa][tablename];
        for (let j = 0; j < tabledata.length; j++) {
            let values = Object.values(tabledata[j]);
            if (tabledata[j].B && tabledata[j].B.includes(yoa + branchcode) || tabledata[j].C && tabledata[j].C.includes(yoa + branchcode)) {
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
    let result = {};
    // let keysk = Object.values(key);
    let values = Object.values(result);

    if (rollno == '') {
        return ['rollno', values];
    }
    const yoa = rollno.substr(0, 4);
    let flag = false;

    for (let i = 1; i <= pages['2023' - yoa]; i++) {
        const tablename = "Table " + i;
        const tableData = data['2023' - yoa][tablename];
        for (let j = 0; j < tableData.length; j++) {
            if (tableData[j].B == rollno || tableData[j].C == rollno || tableData[j].D == rollno) {
                flag = true;
                result = tableData[j];
                break;
            }
        }
        if (flag) break;
    }
    values = Object.values(result);
    return ['rollno', [values]];
}


function getResultByName(name, yoa) {
    let result = [];
    let keysk = Object.values(key);
    let values = Object.values(result);

    if (name == '' || yoa == '') {
        return [];
    }
    let flag = false;

    for (let i = 1; i <= pages['2023' - yoa]; i++) {
        const tablename = "Table " + i;
        const tableData = data['2023' - yoa][tablename];
        for (let j = 0; j < tableData.length; j++) {
            let values = Object.values(tableData[j]);
            for (let k = 0; k < values.length; k++) {
                if (values[k].toString().toUpperCase().includes(name)) {
                    result.push(values);
                    break;
                }
            }
        }
    }
    return ['name', result];
}

// console.log(getResult('2021UIT3137'));

function getRank(grade, branchcode, yoa) {
    if (grade == '' && yoa == '') {
        return 'Rank';
    }
    branchcode = branchcode == '' ? '' : branchcode.toUpperCase();
    let rank = 1;
    for (let i = 1; i <= pages['2023' - yoa]; i++) {
        const tablename = 'Table ' + i;
        const tabledata = data['2023' - yoa][tablename];
        for (let j = 0; j < tabledata.length; j++) {
            let values = Object.values(tabledata[j]);

            if (tabledata[j].B && tabledata[j].B.includes(yoa + branchcode) || tabledata[j].C && tabledata[j].C.includes(yoa + branchcode)) {
                if (typeof (values[values.length - 2]) == 'number') {
                    if (values[values.length - 2] > grade) rank++;
                }
            }
        }
    }
    return rank;
}

// console.log(getRank('9.14', 'UIT', '2023'));

app.get('/getresult', (req, res) => {
    let rolln = '';
    let name = '', yoa = '2021';
    let result = [];
    if (req.query.rollno) {
        rolln = req.query.rollno;
        result = getResult(rolln);
    }
    else if (req.query.name) {
        name = req.query.name;
        yoa = req.query.yoa;
        result = getResultByName(name, yoa);
    }
    // console.log(rolln);
    // console.log(result);
    res.status(200).send(result);
});

app.get('/stats', (req, res) => {
    res.redirect('/getcoursestats');
});

app.get('/result', (req, res) => {
    res.redirect('/getresult');
});

app.get('/getcoursestats', (req, res) => {
    let coursecode = '', branchcode = '', yoa = '';
    if (req.query.coursecode)
        coursecode = req.query.coursecode;
    if (req.query.branchcode)
        branchcode = req.query.branchcode;
    if (req.query.yoa)
        yoa = req.query.yoa;
    const stats = getCourseStats(coursecode, branchcode, yoa);

    res.status(200).send(stats);
});

app.get('/getoverallstats', (req, res) => {
    let branchcode = '', yoa = '';
    if (req.query.branchcode)
        branchcode = req.query.branchcode;
    if (req.query.yoa)
        yoa = req.query.yoa;
    const stats = getOverallStats(branchcode, yoa);
    res.status(200).send(stats);
});

app.get('/getrank', (req, res) => {
    let grade = '', branchcode = '', yoa = '';
    if (req.query.grade)
        grade = req.query.grade;
    if (req.query.branchcode)
        branchcode = req.query.branchcode;
    if (req.query.yoa)
        yoa = req.query.yoa;

    const rank = getRank(grade, branchcode, yoa);
    res.status(200).send(rank.toString());
});

app.listen(PORT, (error) => {
    if (!error) {
        console.log("Server is running on " + PORT);
    }
    else {
        console.log("Error occurred, server can't start", error);
    }
})