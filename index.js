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
const key = {
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

app.get('/getResult', (req, res) => {
    const rolln = req.query.rollno;
    const result = getResult(rolln);
    res.status(200).send(result);
});

app.listen(PORT, (error) => {
    if (!error) {
        console.log("Server is running on " + PORT);
    }
    else {
        console.log("Error occurred, server can't start", error);
    }
})