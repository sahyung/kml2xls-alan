const cheerio = require('cheerio');
const cheerioTableparser = require('cheerio-tableparser');
const fs = require('fs');
const globby = require('globby');
const json2xls = require('json2xls');
const parseKML = require('parse-kml');
const path = require('path');

(async () => {
    const files = await globby('Files', {
        expandDirectories: {
            files: ['*'],
            extensions: ['kml']
        }
    });

    // console.log(files);
    files.forEach(fileInput => {
        let output = fileInput.replace(path.extname(fileInput), '.xlsx');
        // console.log({output});

        fs.access(output, (err) => {
            if (!err) {
                console.log('File', path.basename(output), 'exists. Skip Process!');
                return;
            }

            parseKML
                .toJson(fileInput)
                .then((table) => {
                    // console.log(table.features);
                    let rows = [];
                    table.features.forEach(elem => {
                        if (elem.properties.name !== 'export') {
                            let table = elem.properties.description.replace(/(\t|\n|\r)/gm, '')
                                .split('</B><BR><BR>')
                                .pop()
                                .split('</TABLE>')[0]
                                .concat('</TABLE>');
                            // console.log({table});
                            // console.log('=====');

                            $ = cheerio.load(table);
                            cheerioTableparser($);
                            let data = $("table").parsetable(false, false, true);
                            // console.log({data});

                            let obj = {
                                name: elem.properties.name
                            };
                            data[0].forEach((elem, i) => {
                                if (i) {
                                    obj[elem.slice(0, -1)] = data[1][i];
                                }
                            });
                            rows.push(obj);
                            // console.log({obj});
                        }
                    })
                    // console.log({rows});

                    let xls = json2xls(rows, {
                        fields: {
                            name: 'string',
                            'GPS Week': 'string',
                            'GPS Time': 'string',
                            'Longitude': 'string',
                            'Latitude': 'string',
                            'Height': 'string',
                            'Roll': 'string',
                            'Pitch': 'string',
                            'Yaw': 'string',
                            'Image number': 'string',
                            'Navigation precision': 'string',
                        }
                    });
                    fs.writeFileSync(output, xls, 'binary');
                    console.log('File', path.basename(output), 'Saved!');
                })
                .catch(err => {
                    console.log({
                        err
                    });
                });

        });
    });
})();
