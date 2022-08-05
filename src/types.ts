interface ConvertCsvToXlsxOptionsShared {
    /**
     * Target file name. If left blank, the source file name will be either "<source-file-name>.xlsx" or "<timestamp>.xlsx",
     * depending on whether sourcePath or sourceText was used as source.
     */
    target?: string
    /**
     * Do not write to XLSX file.
     */
    noWrite?: boolean

    /**
     * Set the field delimiter. One character only, defaults to comma (',').
     */
    delimiter?: string
    /**
     * Treat all character after this one as a comment, defaults to '' (disabled).
     */
    comment?: string
}

interface ConvertCsvToXlsxOptionsPath extends ConvertCsvToXlsxOptionsShared {
    sourcePath: string
    sourceText?: never
}
interface ConvertCsvToXlsxOptionsText extends ConvertCsvToXlsxOptionsShared {
    sourcePath?: never
    sourceText: string
}

export type ConvertCsvToXlsxOptions =
    | ConvertCsvToXlsxOptionsPath
    | ConvertCsvToXlsxOptionsText

export function isPathOptions(
    options: ConvertCsvToXlsxOptions
): options is ConvertCsvToXlsxOptionsPath {
    return typeof options.sourcePath !== 'undefined'
}
