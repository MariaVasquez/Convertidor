var fs = require('fs');
const XLSX = require('xlsx');
var nomArchivoTxt = `${__dirname}\\archivostxt`;
var nomArchivoCsv = `${__dirname}\\archivocsv\\parce.csv`;
var nomTest = 'test_data_planning_transport';
function readXlsx() {
    var pathArchivoXlsx = `C:\\Users\\Usuario.SQA-1936\\OneDrive\\ArchivosLaborales`;
    var cont = 1;
    var files = fs.readdirSync(pathArchivoXlsx);
    var nomArchivoXlsx = `${pathArchivoXlsx}\\${nomTest}.xlsx`;
    var file = files.length;
    if (file > 0) {
        for (let i = 0; i < file-1; i++) {
            nomArchivoXlsx = `${pathArchivoXlsx}\\${nomTest} (${cont}).xlsx`;
            try {
                fs.statSync(nomArchivoXlsx);
                cont++;
            }
            catch (err) {
                console.log("No existe archivo xlsx"+nomArchivoXlsx);
            }
        }
    }  
    return nomArchivoXlsx;
}

var workbook = XLSX.readFile(readXlsx(), { cellDates: true });
XLSX.writeFile(workbook, nomArchivoCsv, { bookType: "csv" });

fs.readdir(nomArchivoTxt, (err, files) => {
    fs.readFile(nomArchivoCsv, 'utf8', (err, data) => {
        var dataArray = data.split(/\r?\n/);
        let cont = 1;
        var archivo = `${nomArchivoTxt}\\${nomTest}_${cont}.txt`;
        for (let i = 0; i < files.length + 1; i++) {
            archivo = `${nomArchivoTxt}\\${nomTest}_${cont}.txt`;
            try {
                fs.statSync(archivo);
                cont++;
            }
            catch (err) {
                console.log('Case' + cont);
            }
        }
        var stream = fs.createWriteStream(archivo);
        stream.once('open', () => {
            dataArray.forEach((val, index) => {

                var data = dataArray[index].replace("'", "");
                let includePM = val.includes("PM");
                let includeAM = val.includes("AM");
                if (includePM || includeAM) {
                    data = data.replace(/\s[PA][M]/g, '');
                }
                if (data.includes('n')) {
                    var dataSalt = data.replace(/[,][n]$/g, '');
                    stream.write(dataSalt);
                    stream.write('\r\n');
                }

            });
            stream.end();
        });
    });
});


