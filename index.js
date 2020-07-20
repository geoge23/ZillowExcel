const axios = require('axios')
const enquirer = require('enquirer')
const jsdom = require('jsdom')
const { JSDOM } = jsdom;
const fs = require('fs');
const ExcelJS = require('exceljs');
const NodePath = require('path')
let firstTime = true;
let workbook, worksheet, path;

async function addNew() {    
    const newDoc = {};

    const { url } = await enquirer.prompt({
        type: 'input',
        name: 'url',
        message: 'Enter the Zillow URL 🌐: '
    })

    const { data } = await axios.get(url, {
        headers: {
            'user-agent': 'Mozilla/5.0 (Fortnite; AMD Intel OS 1.2.3) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.1.1 Safari/605.1.15'
        }
    })

    const { window } = new JSDOM(data)
    const doc = window.document
    const list = window.document.querySelector('.ds-home-fact-list');
    const points = list.querySelectorAll('.ds-home-fact-list-item')
    points.forEach(point => {
        const key = point.querySelector('.ds-home-fact-label').textContent;
        const value = point.querySelector('.ds-home-fact-value').textContent;
        switch (key) {
            case 'HOA:':
                newDoc.hoa = value;
                break;

            case 'Cooling:':
                newDoc.cooling = value;
                break;
        
            default:
                break;
        }
    })
    if (newDoc.hoa == undefined) {
        newDoc.hoa = 'None listed'
    }

    const primaryInfo = window.document.querySelector('.ds-summary-row')
    newDoc.price = parseInt(primaryInfo.querySelector('.ds-price').textContent.replace(/\D/g,''))
    const panel = primaryInfo.querySelector('.ds-bed-bath-living-area-container');
    const infos = panel.querySelectorAll('.ds-bed-bath-living-area');

    infos.forEach(info => {
        const splitTuple = info.textContent.split(' ');
        const number = parseInt(splitTuple[0].replace(/\D/g,''))
        switch (splitTuple[1]) {
            case 'bd':
                newDoc.bed = number
                break;
            case 'ba':
                newDoc.bath = number
                break;
            case 'sqft':
                newDoc.sqft = number
                break;
        
            default:
                break;
        }
    })

    const taxTextArray = data.match(/Annual tax amount: <!-- -->\$[0-9]+,[0-9]+/) || data.match(/Tax Annual Amount: <!-- -->[0-9]+\.[0-9]+/)
    const taxText = taxTextArray[0].replace(/\$/g, "");
    console.log(taxText)
    const tax = parseInt(taxText.split(':')[1].split(' <!-- -->')[1].replace(/(,)/, ""));
    newDoc.tax = tax;

    const addressDom = window.document.querySelector('.ds-address-container').childNodes;
    let address = "";
    addressDom.forEach(ad => {
        address += ad.textContent;
    })
    newDoc.address = address;

    try {
        const schoolDom = window.document.querySelector('.ds-agent-listed-schools').childNodes;
        schoolDom.forEach(sch => {
            try {
                const textArray = sch.textContent.split(':')
                newDoc[textArray[0].toLocaleLowerCase()] = textArray[1]
            } catch (_) {}
        })
    } catch (_) {
        console.log('Couldn\'t find schools')
    }

    console.log(newDoc)
    return newDoc
}

async function start() {
    if (firstTime) {
        firstTime = !firstTime;
        const excelPrompt = new enquirer.Select({
            name: 'excelOption',
            message: 'Would you like to',
            choices: ['Import an excel spreadsheet', 'Use or create default sheet in your home directory']
        });
        const choice = await excelPrompt.run()
        switch (choice) {
            case 'Import an excel spreadsheet':
                workbook = new ExcelJS.Workbook()
                const {filePath} = await enquirer.prompt({
                    type: 'input',
                    name: 'filePath',
                    message: 'Input file path to sheet'
                })
                try {
                    await workbook.xlsx.readFile(filePath)
                } catch (e) {
                    console.log('Incorrect path, resetting because' + e.toString())
                    start()
                }
                worksheet = workbook.getWorksheet('Houses')
                path = filePath;
                break;
            case 'Use or create default sheet in your home directory':
                path = NodePath.resolve(require('os').homedir(), 'houses.xlsx')
                workbook = new ExcelJS.Workbook()
                if (!fs.existsSync(path)) {
                    workbook.creator = 'George (me@georgeparks.me)'
                    workbook.lastModifiedBy = 'A literal robot'
                    workbook.created = new Date()
                    worksheet = workbook.addWorksheet('Houses')
                } else {
                    console.log('Using existing workbook...')
                    await workbook.xlsx.readFile(path)
                    worksheet = workbook.getWorksheet('Houses')
                }
                
                break;
        
            default:
                break;
        }

        worksheet.columns = [
            {header: 'Address', key: 'address'},
            {header: 'Price', key: 'price'},
            {header: 'Beds', key: 'bed'},
            {header: 'Bathrooms', key: 'bath'},
            {header: 'Square Feet', key: 'sqft'},
            {header: 'Tax', key: 'tax'},
            {header: 'HOA', key: 'hoa'},
            {header: 'Cooling', key: 'cooling'},
            {header: 'Elementary', key: 'elementary'},
            {header: 'Middle', key: 'middle'},
            {header: 'High', key: 'high'}
        ]

        
        const newDoc = await addNew()
        await worksheet.addRow(newDoc)
        start()
        
    } else {
        const continuePrompt = new enquirer.Select({
            name: 'excelOption',
            message: 'Would you like to',
            choices: ['Add another', 'Save', 'Save and quit', 'Quit']
        });
        const choice = await continuePrompt.run()
        switch (choice) {
            case 'Add another':
                await worksheet.addRow(await addNew())
                start()
                break;
            case 'Save':
                console.log(path)
                await workbook.xlsx.writeFile(path)
                start()
                break;
            case 'Save and quit':
                await workbook.xlsx.writeFile(path)
                return;
            case 'Quit':
                return;
        
            default:
                break;
        }
    }
}

start()
