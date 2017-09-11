'use strict'

var MongoClient = require('mongodb').MongoClient,
    argv = require('yargs').argv,
    ObjectID = require('mongodb').ObjectID,
    co = require('co'),
    test = require('assert'),
    XlsxPopulate = require('xlsx-populate'),
    Promise = require('bluebird'),
    _db,
    _ = require('lodash');


function isPhoto(obj) {
    return _.isString(obj.value) && _.includes(obj.value, '//res.cloudinary.com');
};

function writeToFile(answers) {
    return XlsxPopulate.fromBlankAsync()
        .then(workbook => {
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
};

function forEachSeries(cursor, iterator) {
    return new Promise(function (resolve, reject) {
        var count = 0;
        let _answers = [];
        function processDoc(doc) {
            if (doc != null) {
                count++;
                return iterator(doc).then(function (answers) {

                    _answers.push(...answers);
                    return cursor.next().then(processDoc);
                });
            } else {
                return resolve(_answers);
            }
        }
        cursor.next().then(processDoc);
    });
};

return Promise.resolve(MongoClient.connect(argv.connection))
    .then((db) => {
        _db = db;
        let cursor = db.collection('missiondatas').find({
            'missiondescriptionRef': ObjectID(argv.campaignId)
        });

        return forEachSeries(cursor, function (doc) {

            return Promise.resolve(
                Object.keys(doc)
                .filter(key => _.isPlainObject(doc[key]) && isPhoto(doc[key]))
                .map(key => {
                    let obj = {
                        url: doc[key].value,
                        isCompliant: !doc[key].edit
                    };
                    return obj;
                }));

        });
    })
    .then(writeToFile)
    .catch((err) => console.error("ERROR!!", err))
    .finally (() => {
        if(_db) _db.close();
    });