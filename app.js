const express = require('express')
const multer = require('multer')
const XLSX = require('xlsx')
const fs = require('fs')

const app = express()

const data = require('./constant/data.json')

const storage = multer.diskStorage({
    destination: function (req, file, callback) {
        callback(null, '')
    },
    filename: function (req, file, callback) {
        callback(null, file.originalname)
    },
})

const upload = multer({ storage })

app.post('/import', upload.single('excel'), (req, res) => {


    const file = req.file

    if (file) {
        const work_book = XLSX.readFile(file.path)
        const sheet_name = work_book.SheetNames[0]
        const data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name])

        const result_data = []


        for (let i = 0; i < data.length; i++) {
            const el = data[i];
            result_data.push(el)
        }

        fs.unlinkSync(file.path)
        return res.json({ msg: "Ok", data: result_data })
    }

})

app.get('/export', (req, res) => {
    const ws = XLSX.utils.json_to_sheet(data)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const export_file_name = "export.xlsx";
    XLSX.writeFile(wb, export_file_name);
    const fileStream = fs.createReadStream(
        export_file_name,
    );
    fileStream.on('open', () => {
        res.attachment('export.xlsx');
        fileStream.pipe(res);
        fs.unlinkSync(export_file_name)
    });
    fileStream.on('error', err => {
        next(err);
    });
})

app.listen(9999, () => console.log("Running on port 9999"))