import xlsx = require("xlsx");
import fs = require("fs");
import _ = require("lodash");

export module SpreadSheet {
	export enum EExportReportFieldType {
		String = 0,
		Numeric = 1,
		Date = 2,
		Formula = 3,
	}

	export interface SpreadsheetCell {
		value: string | number | Date;
		type: EExportReportFieldType;
	}

	export interface SpreadsheetRow {
		cells: SpreadsheetCell[];
	}

	export interface Spreadsheet {
		name: string;
		columns: string[];
		data: SpreadsheetRow[];
		createCell?: (item: any) => SpreadsheetCell;
	}

	export function getAddress(row: number, column: number): string  {
		return xlsx.utils.encode_cell({
			c: column,
			r: row,
		});
	}

	export function getRange(fromRow: number, fromColumn: number, toRow: number, toColumn: number): string  {
		return xlsx.utils.encode_range({
			s: {
				c: fromColumn,
				r: fromRow,
			},
			e:{
				c: toColumn,
				r: toRow,
			}
		});
	}

	let excelDateFormat = "yyyy-MM-dd";

	function writeToSheet(spreadsheet: Spreadsheet, worksheet: any) {
		let rowIndex = 0;
		let characterCount = [];

		spreadsheet.columns.forEach(function(name: string, index: number): void {
			characterCount[index] = name.length;
			worksheet[getAddress(rowIndex, index)] = {
				t: "s",
				v: name,
				w: name,
			};
		});

		++rowIndex;

		spreadsheet.data.forEach(function(row: SpreadsheetRow): void {
			_.each(row.cells, function(cellData: SpreadsheetCell, index: number): void {
				if (cellData.value == null)
					return;

				let cell: any = null;
				if (cellData.type == EExportReportFieldType.String) {
					cell = {
						t: "s",
						v: cellData.value,
						w: cellData.value,
					};
				} else if (cellData.type == EExportReportFieldType.Numeric) {
					cell = {
						t: "n",
						v: cellData.value,
					};
				} else if (cellData.type == EExportReportFieldType.Date) {
					cell = {
						t: "d",
						v: <Date>cellData.value,
						z: excelDateFormat,
					};
				} else if (cellData.type == EExportReportFieldType.Formula) {
					cell = {
						t: "n",
						f: cellData.value,
					};
				}

				if (!cell)
					return;

				if (cell.z)
					xlsx.utils.format_cell(cell);

				worksheet[getAddress(rowIndex, index)] = cell;

				if (cell.w)
					characterCount[index] = Math.max(characterCount[index], cell.w.length);
				else if (cell.v)
					characterCount[index] = Math.max(characterCount[index], cell.v.toString().length);

			});

			++rowIndex;
		});

		worksheet["!ref"] = "A1:" + xlsx.utils.encode_cell({
			c: spreadsheet.columns.length - 1,
			r: rowIndex - 1,
		});

		worksheet["!cols"] = _.map(characterCount, function(count: number): any {
			return  {
				wch: count + 4,
			};
		});
	}

	export function writeToFile(spreadsheet: Spreadsheet, path: string): string {
		let worksheet = {};

		// Make sure sheet name is valid

		let invalidCharRegex = /[\\\/\*\[\]\:\?]/g;
		spreadsheet.name = spreadsheet.name.replace(invalidCharRegex, "").substr(0, 31);	// Length restriction on sheet names

		let ampersandCharRegex = /\&/g;
		spreadsheet.name = spreadsheet.name.replace(ampersandCharRegex, "&amp;");	// Must be converted

		if (spreadsheet.name.length == 0)
			spreadsheet.name = "Sheet1";

		writeToSheet(spreadsheet, worksheet);

		let data = "";

		let ssfTable = xlsx.SSF.get_table();
		ssfTable[165] = excelDateFormat;

		let workbook = {
			SheetNames: [spreadsheet.name],
			Sheets: {
				[spreadsheet.name]: worksheet,
			},
			SSF: ssfTable,
		};

		let options: xlsx.WritingOptions = {
			bookType: "xlsx",
			bookSST: false,
			type: "buffer",
			cellDates: true,
		};

		data = xlsx.write(workbook, options);

		fs.writeFileSync(path, data);

		return path;
	}

	export function writeToFileXLSXMutipleSheets(spreadsheets: Spreadsheet[], path: string): string {
		// Make sure sheet name is valid

		let ssfTable = xlsx.SSF.get_table();
		ssfTable[165] = excelDateFormat;

		let workbook = {
			SheetNames: [],
			Sheets: {},
			SSF: ssfTable,
		};

		for (let spreadsheet of spreadsheets) {
			let invalidCharRegex = /[\\\/\*\[\]\:\?]/g;
			spreadsheet.name = spreadsheet.name.replace(invalidCharRegex, "").substr(0, 31);	// Length restriction on sheet names

			let ampersandCharRegex = /\&/g;
			spreadsheet.name = spreadsheet.name.replace(ampersandCharRegex, "&amp;");	// Must be converted

			if (spreadsheet.name.length == 0)
				spreadsheet.name = "Sheet1";

			let worksheet = {};

			writeToSheet(spreadsheet, worksheet);

			workbook.SheetNames.push(spreadsheet.name);
			workbook.Sheets[spreadsheet.name] = worksheet;
		}

		let options: xlsx.WritingOptions = {
			bookType: "xlsx",
			bookSST: false,
			type: "buffer",
			cellDates: true,
		};

		let data = xlsx.write(workbook, options);
		path = path + ".xlsx";
		fs.writeFileSync(path, data);

		return path;
	}
}
