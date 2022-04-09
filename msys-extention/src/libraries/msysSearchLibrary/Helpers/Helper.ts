import { ISearchResult } from "@pnp/sp/search";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { FilterComparisonOperator, IDataFilter } from "@pnp/modern-search-extensibility";
import { isEmpty } from "@microsoft/sp-lodash-subset";
import { stringIsNullOrEmpty } from "@pnp/core";

const LOG_SOURCE: string = 'Helper';
const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const fileExtension = '.xlsx';

export class Helper {
    /**
     * Recupero un array di valori da un array di oggetti
     * @param values 
     * @param propertyName 
     * @returns 
     */
    public static getValuesForArray(values: any[], propertyName: string): any[] {
        let result: any[] = [];
        values.forEach((value, index, array) => {
            result.push(value[propertyName]);
        });

        return result;
    }

    private static formatColumn(worksheet: XLSX.WorkSheet, col: number, t: string, fmt: string) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        // note: range.s.r + 1 skips the header row
        for (let row = range.s.r + 1; row <= range.e.r; ++row) {
            const ref = XLSX.utils.encode_cell({ r: row, c: col });
            if (worksheet[ref] && worksheet[ref].t === 's') {
                worksheet[ref].t = t;
                if (!stringIsNullOrEmpty(fmt)) {
                    worksheet[ref].z = fmt;
                }
            }
        }
    }

    public static downloadExcel(searhResults: ISearchResult[], selectedProperties: string[], headers: string[], types: string[], formats: string[]): void {
        //Elimino le colonne aggiuntive non presenti tra selectedProperties
        var aoo = searhResults;
        if (selectedProperties.length > 0) {
            aoo = searhResults.map((obj) => {
                return selectedProperties.reduce((acc, key) => {
                    acc[key] = obj[key];
                    return acc;
                }, ({}));
            });
        }

        console.log(LOG_SOURCE + " - sheet_data: ", aoo);

        var wb = XLSX.utils.book_new();
        let ws = XLSX.utils.json_to_sheet(aoo);
        console.log(LOG_SOURCE + " - sheet: ", ws);
        //Forzo l'header del file Excel con i valori presi da HEADERS - impostati dall'utente
        ws = XLSX.utils.sheet_add_aoa(ws, [headers], { origin: "A1" });

        // applico la formattazione esempio per currency = '$0.00'
        for (let index = 0; index < types.length; index++) {
            const type = types[index];
            if (!stringIsNullOrEmpty(type) && type !== "String") {
                var t = type == "Date" ? "d" : "n";
                var format = formats[index];
                Helper.formatColumn(ws, index, t, format);
            }
        }

        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        //Genero il BLOB che rappresenta il file XLSX
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const data = new Blob([excelBuffer], { type: fileType });

        //Salvo il file - in questo caso faccio download
        saveAs(data, 'result' + fileExtension);
    }

    /**
   * Build the refinement condition in FQL format
   * @param selectedFilters The selected filter array
   * @param filtersConfiguration The current filters configuration
   * @param moment The moment.js instance to resolve dates
   * @param encodeTokens If true, encodes the taxonomy refinement tokens in UTF-8 to work with GET requests. Javascript encodes natively in UTF-16 by default.
   */
    public static buildFqlRefinementString(selectedFilters: IDataFilter[], moment: any, encodeTokens?: boolean): string[] {

        let refinementQueryConditions: string[] = [];

        selectedFilters.forEach(filter => {

            let operator: any = filter.operator;

            // Mutli values
            if (filter.values.length > 1) {

                let startDate = null;
                let endDate = null;

                // A refiner can have multiple values selected in a multi or mon multi selection scenario
                // The correct operator is determined by the refiner display template according to its behavior
                const conditions = filter.values.map(filterValue => {

                    let value = filterValue.value;

                    if (moment(value, moment.ISO_8601, true).isValid()) {

                        if (!startDate && (filterValue.operator === FilterComparisonOperator.Geq || filterValue.operator === FilterComparisonOperator.Gt)) {
                            startDate = value;
                        }

                        if (!endDate && (filterValue.operator === FilterComparisonOperator.Lt || filterValue.operator === FilterComparisonOperator.Leq)) {
                            endDate = value;
                        }
                    }

                    // If the value is null or undefined, we replace it by the FQL expression string('')
                    // Otherwise the query syntax won't be vaild resuting of to an HTTP 500 
                    if (isEmpty(value)) {
                        value = "string('')";
                    }

                    // Enclose the expression with quotes if the value contains spaces
                    if (/\s/.test(value)) {
                        value = `"${value}"`;
                    }

                    return /ǂǂ/.test(value) && encodeTokens ? encodeURIComponent(value) : value;

                }).filter(c => c);

                if (startDate && endDate) {
                    refinementQueryConditions.push(`${filter.filterName}:range(${startDate},${endDate})`);
                } else {
                    refinementQueryConditions.push(`${filter.filterName}:${operator}(${conditions.join(',')})`);
                }

            } else {

                // Single value
                if (filter.values.length === 1) {

                    const filterValue = filter.values[0];

                    // See https://sharepoint.stackexchange.com/questions/258081/how-to-hex-encode-refiners/258161
                    let refinementToken = /ǂǂ/.test(filterValue.value) && encodeTokens ? encodeURIComponent(filterValue.value) : filterValue.value;

                    // https://docs.microsoft.com/en-us/sharepoint/dev/general-development/fast-query-language-fql-syntax-reference#fql_range_operator
                    if (moment(refinementToken, moment.ISO_8601, true).isValid()) {

                        if (filterValue.operator === FilterComparisonOperator.Gt || filterValue.operator === FilterComparisonOperator.Geq) {
                            refinementToken = `range(${refinementToken},max)`;
                        }

                        // Ex: scenario ('older than a year')
                        if (filterValue.operator === FilterComparisonOperator.Leq || filterValue.operator === FilterComparisonOperator.Lt) {
                            refinementToken = `range(min,${refinementToken})`;
                        }
                    }

                    // If the value is null or undefined, we replace it by the FQL expression string('')
                    // Otherwise the query syntax won't be vaild resuting of to an HTTP 500 
                    if (isEmpty(refinementToken)) {
                        refinementToken = "string('')";
                    }

                    // Enclose the expression with quotes if the value contains spaces
                    if (/\s/.test(refinementToken)) {
                        refinementToken = `"${refinementToken}"`;
                    }

                    refinementQueryConditions.push(`${filter.filterName}:${refinementToken}`);
                }
            }
        });

        return refinementQueryConditions;
    }
}