import { parse as csv } from 'csv-parse'
import xlsx from 'xlsx'
import { ConvertCsvToXlsxOptions, isPathOptions } from './types'
import fs from 'fs'
import path from 'path'

function getSource(options: ConvertCsvToXlsxOptions) {
    if (!isPathOptions(options)) return options.sourceText

    if (!fs.existsSync(options.sourcePath))
        throw new Error(`Source file not found: ${options.sourcePath}`)

    return fs.readFileSync(options.sourcePath, {
        encoding: 'utf8',
        flag: 'r',
    })
}

function getTargetFileName(options: ConvertCsvToXlsxOptions) {
    if (!isPathOptions(options)) return `${new Date().toISOString()}.xlsx`
    const fileName = options.sourcePath.replace(/\.[^/.]+$/, '.xlsx')
    return path.basename(fileName)
}

function validateOptions(options: ConvertCsvToXlsxOptions) {
    if (options.delimiter && options.delimiter.length !== 1)
        throw new Error(`Invalid delimiter: ${options.delimiter}`)
}

function parseCSV(options: ConvertCsvToXlsxOptions) {
    const source = getSource(options)

    return new Promise<string[][]>((resolve, reject) => {
        const parser = csv({
            delimiter: options.delimiter || ',',
            comment: options.comment || '#',
            relaxQuotes: true,
            ltrim: true,
            rtrim: true,
        })

        const records: string[][] = []

        parser.on('readable', () => {
            let record: string[]
            while ((record = parser.read()) !== null) records.push(record)
        })

        parser.on('error', err => {
            reject(err.message)
        })

        parser.on('end', () => {
            resolve(records)
        })

        parser.write(source)
        parser.end()
    })
}

function createXlsx(records: string[][]) {
    const book = xlsx.utils.book_new()
    const sheet = xlsx.utils.aoa_to_sheet(records)
    xlsx.utils.book_append_sheet(book, sheet)

    return book
}

export default async function convertCsvToXlsx(
    options: ConvertCsvToXlsxOptions
) {
    validateOptions(options)

    const records = await parseCSV(options)
    const book = createXlsx(records)

    const outFile = getTargetFileName(options)
    xlsx.writeFile(book, outFile)

    return book
}
