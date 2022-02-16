import { ISearchResult } from "@pnp/sp/search";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

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

    public static downloadExcel(searhResults: ISearchResult[], selectedProperties: string[], headers: string[]): void {
        //Elimino le colonne aggiuntive non presenti tra selectedProperties
        var aoo = searhResults.map((obj) => {
            return selectedProperties.reduce((acc, key) => {
                acc[key] = obj[key];
                return acc;
            }, ({}));
        });
        console.log(LOG_SOURCE + " - sheet_data: ", aoo);

        var wb = XLSX.utils.book_new();
        //let ws = XLSX.utils.json_to_sheet(searhResults);
        let ws = XLSX.utils.json_to_sheet(aoo);
        console.log(LOG_SOURCE + " - sheet: ", ws);
        //Forzo l'header del file Excel con i valori presi da HEADERS - impostati dall'utente
        ws = XLSX.utils.sheet_add_aoa(ws, [headers], { origin: "A1" });
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        //Genero il BLOB che rappresenta il file XLSX
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const data = new Blob([excelBuffer], { type: fileType });

        //Salvo il file - in questo caso faccio download
        saveAs(data, 'result' + fileExtension);
    }
}