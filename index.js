require('dotenv').config();
const argv = require('minimist')(process.argv.slice(2));

const fs = require('fs');
const path = require('path');
const xlsxPopulate = require('xlsx-populate');

const moment = require('moment');
const MyImap = require('./my-imap');
const logger = require('pino')({
    transport: {
        target: 'pino-pretty',
        options: {
            translateTime: false,
            colorize: true,
            ignore: 'pid,hostname,time',
        },
    },
});

async function convertXlsxToCsv(xlsxPath, csvPath) {
    const workbook = await xlsxPopulate.fromFileAsync(xlsxPath);
    const csvData = workbook.sheet(0).usedRange().value().map(row => row.join(',')).join('\n');
    fs.writeFileSync(csvPath, csvData);
}

async function run(subject) {
    const config = {
        imap: {
            user: process.env.EMAIL_USER,
            password: process.env.EMAIL_PASSWORD,
            host: process.env.EMAIL_HOST,
            port: process.env.EMAIL_PORT,
            tls: process.env.EMAIL_TLS,
        },
        debug: logger.info.bind(logger),
    };

    const imap = new MyImap(config);
    const result = await imap.connect();
    logger.info(`result: ${result}`);
    const boxName = await imap.openBox();
    logger.info(`boxName: ${boxName}`);

    const criteria = [];
    criteria.push('UNSEEN');
    criteria.push(['SINCE', moment().format('MMMM DD, YYYY')]);
    if (subject) {
        criteria.push(['HEADER', 'SUBJECT', subject]);
    }

    const emails = await imap.fetchEmails(criteria);

    const reportsDir = path.join(__dirname, 'reports');
    if (!fs.existsSync(reportsDir)) {
        fs.mkdirSync(reportsDir);
    }

    for (const email of emails) {
        for (const file of email.files) {
            const ext = path.extname(file.originalname);
            if (ext === '.xlsx') {
                const timestamp = moment().format('YYYYMMDD_HHmmss');
                const filename = path.basename(file.originalname, ext);
                const newFilename = `${filename}_${timestamp}${ext}`;
                const destinationPath = path.join(__dirname, 'reports', newFilename);
                const csvPath = path.join(__dirname, 'reports', `${filename}_${timestamp}.csv`);
                try {
                    fs.writeFileSync(destinationPath, file.buffer);
                    await convertXlsxToCsv(destinationPath, csvPath);
                    logger.info(`Converted ${newFilename} to CSV`);
                } catch (error) {
                    logger.error(`Error while writing or converting file: ${error}`);
                }
            }
        }
    }
}

run(argv.subject).then(() => {
    process.exit();
}).catch((error) => {
    logger.error(error);
    process.exit(1);
});
