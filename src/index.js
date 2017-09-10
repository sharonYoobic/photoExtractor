'use strict'

var MongoClient = require('mongodb').MongoClient,
    argv = require('yargs').argv,
    ObjectID = require('mongodb').ObjectID,
    co = require('co'),
    test = require('assert'),
    XlsxPopulate = require('xlsx-populate'),
    _ = require('lodash');

co(function* () {
    let db;
    try {

        db = yield MongoClient.connect(argv.connection);

        let cursor = db.collection('missiondatas').find({
            'missiondescriptionRef': ObjectID(argv.campaignId)
        });

        let answers = [];
        while (yield cursor.hasNext()) {
            let doc = yield cursor.next();
            Object.keys(doc)
                .filter(key => {
                    return _.isPlainObject(doc[key]) && _.isString(doc[key].value) && doc[key].value.includes('//res.cloudinary.com')
                })
                .map(key => {
                    return {
                        url: doc[key].value,
                        isCompliant: !doc[key].edit
                    };
                })
                .forEach(answer => answers.push(answer));

        }
        writeToFile(answers);

    } catch (error) {
        console.log("ERROR!!", error);
    } finally {
        db.close();
    }
});

function writeToFile(answers) {
    XlsxPopulate.fromBlankAsync().then(workbook => {
        // Modify the workbook.
        workbook.sheet(0).row(1).cell(1).value("url");
        workbook.sheet(0).row(1).cell(2).value("isCompliant");

        answers.forEach((answer, index) => {
            workbook.sheet(0).row(index + 2).cell(1).value(answer.url);
            workbook.sheet(0).row(index + 2).cell(2).value(answer.isCompliant);
        });

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
}