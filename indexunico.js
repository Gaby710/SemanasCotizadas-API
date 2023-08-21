//  -   -   -   -   -   -   -   -   -   -   N O D E   M O D U L E S  -   -   -   -   -   -   -   -   -   -
var express = require('express');
var bodyParser = require('body-parser');
var dotenv = require('dotenv');
var nodemailer = require('nodemailer');
var fs = require('fs');
var webdriver = require('selenium-webdriver'),
    Builder = webdriver.Builder,
    By = webdriver.By,
    Key = webdriver.Key,
    until = webdriver.until;
const chrome = require('selenium-webdriver/chrome');
var imaps = require('imap-simple');
var pdf2base64 = require('pdf-to-base64');
const pdf = require('pdf-parse');
const request = require('request');
var mysql = require('mysql');
const sleep = require('sleep-promise');
const htmlToText = require('html-to-text');
const { v4: uuidv4 } = require('uuid');
dotenv.config();

const proxyListImss = ['92.118.40.82:8800', '92.118.40.149:8800', '92.118.40.253:8800', '92.118.40.95:8800', '92.118.40.187:8800', '92.118.40.193:8800', '69.147.248.35:8800', '69.147.248.217:8800', '92.118.40.30:8800', '92.118.40.222:8800'];

//  -   -   -   -   -   -   -   M A N E J O - L O G S -   -   -   -   -   -  -   -  -  -

Object.defineProperty(global, '__stack', {
    get: function () {
        var orig = Error.prepareStackTrace;
        Error.prepareStackTrace = function (_, stack) {
            return stack;
        };
        var err = new Error();
        Error.captureStackTrace(err, arguments.callee);
        var stack = err.stack;
        Error.prepareStackTrace = orig;
        return stack;
    },
});

Object.defineProperty(global, '__line', {
    get: function () {
        return __stack[1].getLineNumber();
    },
});

Object.defineProperty(global, '__function', {
    get: function () {
        return __stack[1].getFunctionName();
    },
});

//  -   -   -   -   -   -   -   E R R O R - NO - T R A T A D O -   -   -   -   -   -  -   -  -  -

process.on('uncaughtException', (err) => {
    // Cuando se presente un error, ejecuta esta última función
    notificaError(err, function (error, data) {
        if (error) console.error(error);

        console.log(data);
        console.error('Se envió notificación de la falla: ' + err);
    });
});

//  -   -   -   -   -   -   -   -   -   -   C O N S T A N T E S -   -   -   -   -   -   -   -   -   -
const browserSearch = 'chrome';

const pool = mysql.createPool({
    connectionLimit: 15,
    host: process.env.MYSQL_HOST,
    user: process.env.MYSQL_USER,
    password: process.env.MYSQL_PASS,
    database: process.env.MYSQL_DB,
    port: 3306,
});

//  -   -   -   -   -   -   -   -   -   -   C O N E X I Ó N   A P I    -   -   -   -   -   -   -   -   -

const PORT = process.env.PORT || 9822;
const app = express();
app.set('port', PORT);
const server = app.listen(app.get('port'));
app.use(bodyParser.json());
console.log("I'm wating for you @" + PORT);
server.timeout = 240000;

//  -   -   -   -   -   -   -   -   -   -   F  L U J O   -   -   -   -   -   -   -   -   -   -

app.post('/busquedaRPA', async function (req, res) {
    console.log(JSON.stringify(req.body));
    let respuesta = {};
    var infoUsuario = req.body;
    let idUnico = uuidv4();
    infoUsuario.idUnico = idUnico;

    try {
        await guardaUsuario(await generaFechaSQL(), infoUsuario.token, infoUsuario.nss, infoUsuario.curp, infoUsuario.idUnico);
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        res.status(500).send({ message: 'Ha ocurrido un error: ' + error.message, status: 500 });
    }

    if (infoUsuario.nss.length != 11 || infoUsuario.curp.length != 18) {
        res.status(400).send({ message: 'Los parametros enviados no son correctos', status: 400 });
    } else {
        res.status(200).send({ message: 'Espere mientras su petición es atendida.', id: idUnico });
        try {
            respuesta = await extraccionDatos(infoUsuario);
        } catch (error) {
            console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        }
    }
});

app.post('/obtenerResultado', async function (req, res) {
    try {
        var infoUsuario = req.body;
        let idPeticion = infoUsuario.id;
        let informacion = await buscaInformacion(idPeticion);
        console.log('Se recibe petición de búsqueda:' + idPeticion);
        if (informacion.length == 0) {
            console.log('Petición no encontrada');
            res.status(200).send('Esta petición no existe, intente con otro identificador');
        } else if (informacion[0].statusBusqueda == null) {
            console.log('Petición encontrada no finalizada');
            res.status(200).send('Esta petición aún no ha terminado de procesarse, favor de esperar');
        } else {
            console.log('Petición encontrada finalizada');
            delete informacion[0].statusBusqueda;
            let archivo;
            if (fs.existsSync(process.env.pdfPath + informacion[0].nss + '_SemanasCotizadas.pdf')) {
                archivo = await pdf2base64(process.env.pdfPath + informacion[0].nss + '_SemanasCotizadas.pdf');
            } else {
                archivo = 'Archivo no disponible';
            }
            informacion[0].archivo = archivo;
            res.status(200).send(informacion[0]);
        }
    } catch (error) {
        res.status(500).send('Ha ocurrido un error: ' + error.message);
    }
});

//  -   -   -   -   -   -   -   -   -   -   F U N C I O N E S   -   -   -   -   -   -   -   -   -

async function guardaUsuario(fecha, token, nss, curp, idUnico) {
    return new Promise((resolve, reject) => {
        pool.getConnection((error, connection) => {
            if (error) {
                console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                resolve(0);
                return;
            }
            const query = 'INSERT INTO registros_SemanasCotizadas(fecha,token,nss,curp,idRPA) VALUES (?,?,?,?,?)';

            connection.query(query, [fecha, token, nss, curp, idUnico], (error, results) => {
                connection.release();

                if (error) {
                    console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                    resolve(0);
                    return;
                }

                resolve();
                return;
            });
        });
    });
}

async function guardaNombre(idUnico, nombre) {
    return new Promise((resolve, reject) => {
        pool.getConnection((error, connection) => {
            if (error) {
                console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                resolve(0);
                return;
            }

            const query = 'UPDATE registros_SemanasCotizadas SET nombre=? where idRPA =?;';

            connection.query(query, [nombre, idUnico], (error, results) => {
                connection.release();

                if (error) {
                    console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                    resolve(0);
                    return;
                }

                resolve();
                return;
            });
        });
    });
}

async function guardaDescripcionError(idUnico, descripcion) {
    return new Promise((resolve, reject) => {
        pool.getConnection((error, connection) => {
            if (error) {
                console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                resolve(0);
                return;
            }

            const query = 'UPDATE registros_SemanasCotizadas SET descError=? where idRPA =?;';

            connection.query(query, [descripcion, idUnico], (error, results) => {
                connection.release();

                if (error) {
                    console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                    resolve(0);
                    return;
                }
                resolve();
                return;
            });
        });
    });
}

async function guardaDatosExtraidos(idUnico, datos) {
    return new Promise((resolve, reject) => {
        pool.getConnection((error, connection) => {
            if (error) {
                console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                resolve(0);
                return;
            }
            const query = 'UPDATE registros_SemanasCotizadas SET fechaAlta = ?, fechaBaja = ?, salarioBase  = ?, ultimoPatron = ?, ultimoRegistroPatronal = ?, nombreArchivo = ? where idRPA=?;';

            connection.query(query, [datos.fechaAlta, datos.fechaBaja, datos.salarioBase, datos.ultimoPatron, datos.ultimoRegistroPatronal, datos.nombreArchivo, idUnico], (error, results) => {
                connection.release();

                if (error) {
                    console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                    resolve(0);
                    return;
                }

                resolve();
                return;
            });
        });
    });
}

async function guardaStatus(idUnico, status) {
    return new Promise((resolve, reject) => {
        pool.getConnection((error, connection) => {
            if (error) {
                console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                resolve(0);
                return;
            }

            const query = 'UPDATE registros_SemanasCotizadas SET statusBusqueda=? where idRPA =?;';

            connection.query(query, [status, idUnico], (error, results) => {
                connection.release();

                if (error) {
                    console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                    resolve(0);
                    return;
                }

                resolve();
                return;
            });
        });
    });
}

async function buscaInformacion(idUnico) {
    return new Promise((resolve, reject) => {
        pool.getConnection((error, connection) => {
            if (error) {
                console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                resolve([]);
                return;
            }
            const query = 'select nss,curp,nombre,fechaBaja,fechaAlta,nombre,salarioBase,ultimoPatron,ultimoRegistroPatronal,descError,statusBusqueda,nombreArchivo from registros_SemanasCotizadas where idRPA =?;';
            connection.query(query, [idUnico], (error, results) => {
                connection.release();

                if (error) {
                    console.error(`Funcion: ${__function} - Linea: ${__line} - Error: ${error.stack || error}`);
                    resolve([]);
                    return;
                }

                resolve(JSON.parse(JSON.stringify(results)));
                return;
            });
        });
    });
}

async function generaFechaSQL() {
    var fechaSql = new Date();
    var horario = fechaSql.getTimezoneOffset() / 60;

    //fechaSql.setHours(fechaSql.getHours() - horario);
    fechaSql.setHours(fechaSql.getHours());
    return fechaSql;
}

async function extraccionDatos(usuario, intento = 0) {
    let respuestaFlujo;

    if (intento < process.env.intentosRpa) {
        console.log(`Comenzando intento: ${intento}`);
        respuestaFlujo = await flujoRPA(usuario, intento);
        if (respuestaFlujo.status !== 100 && respuestaFlujo.status !== 402) {
            intento = intento + 1;
            console.log(`Se realiza nuevamente el flujo del RPA`);
            return await extraccionDatos(usuario, intento);
        }
    }

    console.log('FIN INTENTOS');

    if (!respuestaFlujo) {
        respuestaFlujo = {};
        respuestaFlujo.status = 402;
    }

    await guardaStatus(usuario.idUnico, respuestaFlujo.status);
    return respuestaFlujo;
}

async function flujoRPA(usuario, contador) {
    var CURP = usuario.curp;
    var NSS = usuario.nss;
    var CORREO = process.env.correoImss;
    var NombreCompleto = '';
    var idUnico = usuario.idUnico;
    let objresp = {};
    objresp.curp = CURP;
    objresp.nss = NSS;

    try {
        const number = Math.floor(Math.random() * proxyListImss.length);
        const proxySelected = proxyListImss[number]
        console.log(`Proxy seleccionado para semanas cotizadas: ${proxySelected}`)
        let opts = new chrome.Options().addArguments(`--proxy-server=http://${proxySelected}`);
        opts.addArguments("--ignore-certificate-errors");
        var driver = new webdriver.Builder().forBrowser(browserSearch).setChromeOptions(opts).build();
        //var driver = new webdriver.Builder().forBrowser(browserSearch).build();
        //await driver.manage().window().setRect({ height: 900, width: 1366, x: 2000, y: 0 })
        // await driver.manage().window().setRect({height: 900, width: 1366, x:0, y:0})
        await driver.get(process.env.pagSemanasCotizadas);
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        driver.close();
        objresp.status = 301;
        return objresp;
    }

    try {
        await sleep(3000);
        const elemento = await driver.wait(until.elementLocated(By.xpath('//*[@id="captchaImg"]')), 6000);
        const actions = driver.actions({ async: true });
        await actions.move({ origin: elemento }).perform();
        var image = await elemento.takeScreenshot();
        var response = await captchaResolve(image);
        var captchaText = response;
        console.log(`Linea ${__line} - CAP1: ${captchaText}`);
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        driver.close();
        objresp.status = 201;
        return objresp;
    }

    try {
        await driver.findElement(By.xpath('//*[@id="captcha"]'), 6000).sendKeys(captchaText);
        await driver.findElement(By.xpath('//*[@id="CURP"]'), 6000).sendKeys(CURP);
        await driver.findElement(By.xpath('//*[@id="NSS"]'), 6000).sendKeys(NSS);
        await driver.findElement(By.xpath('//*[@id="Correo"]'), 6000).sendKeys(CORREO);
        await driver.executeScript('window.scrollBy(0,100)');

        await driver.wait(until.elementLocated(By.xpath('//*[@id="btnContinuar"]')), 6000).click();
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        driver.close();
        objresp.status = 201;
        return objresp;
    }

    try {
        console.log(`Verificamos existencia de alerta al llenar formulario`);

        var textAlert = await driver.wait(until.elementLocated(By.xpath('//*[@id="divErrorCampos"]/p')), 6000).getText();
        if (textAlert != 'La información capturada no coincide con la imagen presentada.') {
            console.log(`No se pudo ingresar debido al siguiente error: ${textAlert}`);
            await guardaDescripcionError(idUnico, textAlert);
            driver.close();
            objresp.status = 402;
            objresp.descError = textAlert;
            return objresp;
        } else {
            driver.close();
            objresp.status = 201;
            return objresp;
        }
    } catch (error) {
        console.log(`No hubo alertas al pasar el formulario del imss`);
    }

    try {
        console.log(`Verificamos existencia de error al llenar formulario`);

        var textError = await driver.wait(until.elementLocated(By.xpath('//*[@id="mensajesError"]/div/p')), 6000).getText();

        console.log(`No se pudo ingresar debido al siguiente error: ${textError}`);
        await guardaDescripcionError(idUnico, textError);
        driver.close();
        objresp.status = 402;
        objresp.descError = textError;
        return objresp;
    } catch (error) {
        console.log(`No hubo error al pasar el formulario del imss`);
    }

    try {
        await driver.wait(until.elementLocated(By.xpath('//*[@id="formTurnar"]/div[2]/div/div[1]/h4/button')), 6000);
        NombreCompleto = await driver.wait(until.elementLocated(By.xpath('//*[@id="formTurnar"]/div/div[2]/div/table/tbody/tr/td/table/tbody/tr[3]/td/span')), 5000).getText();
        await guardaNombre(idUnico, NombreCompleto);
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        driver.close();
        objresp.status = 201;
        return objresp;
    }

    try {
        var codigoSegundoCaptcha = await segundoCaptchaRecursivo(driver);

        if (codigoSegundoCaptcha != 100) {
            console.log(`${__function} - Linea ${__line} - Error en resolver el segundo captcha con codigo ${codigoSegundoCaptcha}`);
            driver.close();
            return codigoSegundoCaptcha;
        }
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        driver.close();
        objresp.status = 301;
        return objresp;
    }

    try {
        var downloadText = await DescagarTextoEmail(NombreCompleto);

        if (downloadText == 202) {
            var windows = await driver.getAllWindowHandles();
            if (windows.length > 1) {
                await driver.switchTo().window(windows[1]);
                await sleep(500);
                driver.close();
                await driver.switchTo().window(windows[0]);
                await sleep(500);
            }
            driver.close();
            objresp.status = 202;
            return objresp;
        }

        var urlConfirm = await analizeTextHtmlEmail(downloadText);
        if (urlConfirm == 203) {
            var windows = await driver.getAllWindowHandles();
            if (windows.length > 1) {
                await driver.switchTo().window(windows[1]);
                await sleep(500);
                driver.close();
                await driver.switchTo().window(windows[0]);
                await sleep(500);
            }
            driver.close();
            objresp.status = 203;
            return objresp;
        }
        // console.log("---------------------")
        // console.log(urlConfirm)
        // console.log("---------------------")
        newPageLink = "window.open('" + urlConfirm + "','_blank')";
        await driver.executeScript(newPageLink);
        await sleep(6000);
        var windows = await driver.getAllWindowHandles();
        await driver.switchTo().window(windows[1]);
        await sleep(1000);
        driver.close();
        await driver.switchTo().window(windows[0]);
        await sleep(1000);
        driver.close();
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        var windows = await driver.getAllWindowHandles();
        if (windows.length > 1) {
            await driver.switchTo().window(windows[1]);
            driver.close();
            await driver.switchTo().window(windows[0]);
        }
        driver.close();
        objresp.status = 301;
        return objresp;
    }

    try {
        var respuesta = await DescagarPdfEmail(NSS);

        if (respuesta == 204) {
            objresp.status = 204;
            return objresp;
        }

        var textoPdf = await getPDFtext(NSS);

        if (textoPdf == 205) {
            objresp.status = 205;
            return objresp;
        }

        var datosPdf = await analizaTextoPDF(textoPdf);

        if (datosPdf == 205) {
            objresp.status = 205;
            return objresp;
        }

        console.log(datosPdf);
        datosPdf.nombreArchivo = NSS + '_SemanasCotizadas.pdf';
        await guardaDatosExtraidos(idUnico, datosPdf);

        var archivo64 = await pdf2base64(process.env.pdfPath + NSS + '_SemanasCotizadas.pdf');

        objresp.fechaBaja = datosPdf.fechaBaja;
        objresp.fechaAlta = datosPdf.fechaAlta;
        objresp.salarioBase = datosPdf.salarioBase;
        objresp.ultimoPatron = datosPdf.ultimoPatron;
        objresp.ultimoRegistroPatronal = datosPdf.ultimoRegistroPatronal;
        objresp.archivo64 = archivo64;
        objresp.nombreArchivo = NSS + '_SemanasCotizadas.pdf';
        objresp.status = 100;

        // try {
        //     fs.unlinkSync(process.env.pdfPath + NSS + "_SemanasCotizadas.pdf")
        //     console.log("Archivo de semanas cotizadas eliminado")
        // } catch (e) {
        //     console.log("No se ha podido eliminar el archivo de semanas cotizadas.")
        // }

        return objresp;
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        objresp.status = 205;
        return objresp;
    }
}

async function captchaResolve(image) {
    try {
        var response = await resolve_AC(image);
        return response;
    } catch (error) {
        // console.log("F captchaResolve 1- Ocurrio un problema al resolver el captcha: " + error)
        console.error(`${__function} - Linea ${__line} - Ocurrio un problema al resolver el captcha: ${error}`);
        return '00000';
    }
}

async function resolve_AC(image) {
    return new Promise((resolve, reject) => {
        var key = process.env.antiCaptchaKey;
        var timeout = 25000;

        request.post({ url: 'http://anti-captcha.com/in.php', form: { method: 'base64', key: key, body: image } }, function (error, response, body) {
            if (error) {
                console.error(`${__function} - Linea ${__line} - Error: ${error}`);
                reject(error);
                return;
            }

            setTimeout(() => {
                request.get('http://anti-captcha.com/res.php?action=get&id=' + body.split('|')[1] + '&key=' + key, function (error, response, body) {
                    if (error) {
                        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
                        reject(error);
                        return;
                    }

                    var captcharesult = body.split('|')[1];

                    if (captcharesult == undefined || captcharesult == null || captcharesult == '') {
                        resolve('000000');
                    } else {
                        resolve(captcharesult);
                    }
                });
            }, timeout);
        });
    });
}

async function DescagarPdfEmail(NSS) {
    await sleep(process.env.tiempoEspera);

    var config = {
        imap: {
            user: process.env.Correo_Imss,
            password: process.env.passCorreo_Imss,
            host: 'imap.gmail.com',
            port: 993,
            tls: true,
            tlsOptions: { servername: 'imap.gmail.com' },
            authTimeout: 3000,
        },
    };

    try {
        var connection = await imaps.connect(config);
        await connection.openBox('INBOX');
        var delay = 24 * 3600 * 1000;
        var yesterday = new Date();
        yesterday.setTime(Date.now() - delay);
        yesterday = yesterday.toISOString();

        var searchCriteria = ['ALL', ['BODY', NSS], ['SINCE', yesterday]];
        var fetchOptions = { bodies: ['HEADER.FIELDS (FROM TO SUBJECT DATE)'], struct: true };

        var messages = await connection.search(searchCriteria, fetchOptions);
        var attachments = [];

        messages.forEach(function (message) {
            var parts = imaps.getParts(message.attributes.struct);

            attachments = attachments.concat(
                parts
                    .filter(function (part) {
                        return part.disposition && part.disposition.type.toUpperCase() === 'ATTACHMENT';
                    })
                    .map(function (part) {
                        // retrieve the attachments only of the messages with attachments
                        return connection.getPartData(message, part).then(function (partData) {
                            return {
                                filename: part.disposition.params.filename,
                                data: partData,
                            };
                        });
                    })
            );
        });

        var attachmentsDownload = await Promise.all(attachments);

        attachmentsDownload.forEach((documentos) => {
            fs.appendFileSync(process.env.pdfPath + NSS + '_SemanasCotizadas.pdf', new Buffer.from(documentos.data));
        });

        console.log('Pdf descargado');
        connection.end();
        return 'OK';
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        notificaError(`${__function} - Linea ${__line} - Error: ${error}`);
        return 204;
    }
}

async function getPDFtext(NSS) {
    try {
        var dataBuffer = fs.readFileSync(process.env.pdfPath + NSS + '_SemanasCotizadas.pdf');
        var data = await pdf(dataBuffer);
        return data.text.replace(/(?:\\[rn]|[\r\n]+)+/g, '<Enter>');
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        return 205;
    }
}

async function analizaTextoPDF(texto) {
    try {
        let salario = texto.match(/(?<=\<Enter\>\$ )[0-9.]+(?=\<Enter\>)/gi);
        let nombrePatron = texto.match(/(?<=nombre del patrón)[\s\w,]+(?=\<Enter\>)/gi);
        let registroPatron = texto.match(/(?<=Registro Patronal)[\s\w]+(?=\<Enter\>)/gi);
        let fechas = texto.match(/[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2,4}|Vigente/gi);

        var datosUsuario = {
            fechaBaja: '',
            fechaAlta: '',
            salarioBase: '',
            ultimoPatron: '',
            ultimoRegistroPatronal: '',
        };

        if (nombrePatron.length > 0) {
            datosUsuario.fechaBaja = fechas[0];
            datosUsuario.fechaAlta = fechas[1];
            datosUsuario.salarioBase = salario[0];
            datosUsuario.ultimoPatron = nombrePatron[0];
            datosUsuario.ultimoRegistroPatronal = registroPatron[0];
        }

        return datosUsuario;
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        return 205;
    }
}

async function DescagarTextoEmail(nombre) {
    await sleep(process.env.tiempoEspera);

    var config = {
        imap: {
            user: process.env.Correo_Imss,
            password: process.env.passCorreo_Imss,
            host: 'imap.gmail.com',
            port: 993,
            tls: true,
            tlsOptions: { servername: 'imap.gmail.com' },
            authTimeout: 3000,
        },
    };

    try {
        var connection = await imaps.connect(config);
        await connection.openBox('INBOX');
        var delay = 24 * 3600 * 1000;
        var yesterday = new Date();
        yesterday.setTime(Date.now() - delay);
        yesterday = yesterday.toISOString();

        var searchCriteria = ['ALL', ['BODY', nombre], ['SINCE', yesterday], ['SUBJECT', 'Servicio Digital: Solicitud de Constancia de Semanas Cotizadas del Asegurado']];
        var fetchOptions = { bodies: ['HEADER', 'TEXT'], struct: true };

        var results = await connection.search(searchCriteria, fetchOptions);
        var subjects = results.map(function (res) {
            return res.parts.filter(function (part) {
                return part.which === 'TEXT';
            })[0].body;
        });

        connection.end();

        var lastSubject = subjects.pop();
        // console.log(lastSubject);

        return lastSubject != null && lastSubject != undefined ? lastSubject : 202;
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        notificaError(`${__function} - Linea ${__line} - Error: ${error}`);
        return 202;
    }
}

async function analizeTextHtmlEmail(text) {
    try {
        var text1 = `<html>${text.split('<html>')[1].split('</html>')[0]}</html>`;
        var text2 = text1.replace(/=\r\n/g, '');
        text2 = text2.replace(/=3D/g, '=');
        var text2Dec = htmlToText.fromString(text2, {});

        var urlConfirm = text2Dec.split('Solicitud de Constancia de Semanas Cotizadas del Asegurado')[1].split('Si no solicitaste esto, ignora este correo')[0];
        urlConfirm = urlConfirm.split('[')[1].split(']')[0];

        return urlConfirm;
    } catch (error) {
        console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        return 203;
    }
}

async function segundoCaptchaRecursivo(driver, codigo = 100, intento = 0) {
    console.log(`${__function} --- intento: ${intento} --- codigo: ${codigo}`);

    if (intento < process.env.intentosRpa) {
        try {
            await driver.wait(until.elementLocated(By.xpath('//*[@id="formTurnar"]/div[2]/div/div[1]/h4/button')), 10000).click();
            await driver.wait(until.elementLocated(By.xpath('//*[@id="detalle"]')), 15000).click();
        } catch (error) {
            console.error(`${__function} - Linea ${__line} - Error: ${error}`);
            return await segundoCaptchaRecursivo(driver, 301, intento + 1);
        }

        try {
            const elemento = await driver.wait(until.elementLocated(By.xpath('//*[@id="captchaImg"]')), 6000);
            const actions = driver.actions({ async: true });
            await actions.move({ origin: elemento }).perform();
            var image = await elemento.takeScreenshot();
            var response = await captchaResolve(image);
            var captchaText = response;
            console.log(`Linea ${__line} - CAP2: ${captchaText}`);
        } catch (error) {
            console.error(`${__function} - Linea ${__line} - Error: ${error}`);
            return await segundoCaptchaRecursivo(driver, 201, intento + 1);
        }

        try {
            await driver.findElement(By.xpath('//*[@id="captcha"]'), 6000).sendKeys(captchaText);
            await driver.wait(until.elementLocated(By.xpath('//*[@id="btnContinuar"]')), 6000).click();
        } catch (error) {
            console.error(`${__function} - Linea ${__line} - Error: ${error}`);
            return await segundoCaptchaRecursivo(driver, 201, intento + 1);
        }

        try {
            await driver.wait(until.elementLocated(By.xpath('//*[@id="btnCerrar"]')), 6000).click();
            console.log(`${__function} - Linea ${__line} - No ingreso en el segundo captcha`);
            return await segundoCaptchaRecursivo(driver, 201, intento + 1);
        } catch (error) {
            console.log('Consulta Finalizada con Exito');
            // await sleep(15000);
            codigo = 100;
        }
    }

    return codigo;
}

//  -   -   -   -   -   -   -   -   -   -   N O T I F I C A R - E R R O R   -   -   -   -   -   -   -   -   -

async function notificaError(err) {
    let transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: 'alertasxira@gmail.com',
            pass: 'xira2018',
        },
        tls: {
            rejectUnauthorized: false,
        },
    });

    let mailOptions = {
        from: 'infobot@xira-intelligence.com',
        to: process.env.correoError,
        subject: 'FALLA RPA PENTAFON SEMANAS COTIZADAS - NOMINA!!',
        text: '',
        html: 'El robot presentó una falla => ' + err,
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.error(`${__function} - Linea ${__line} - Error: ${error}`);
        } else {
            console.log(info);
        }
    });
    return 'OK';
}
