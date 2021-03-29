import * as fs from "fs";
import * as path from "path";
import * as util from "util";
import * as stream from 'stream';

import xlsx = require("xlsx");
import https = require("https");
import url = require("url");
import moment = require("moment");
import sanitizeFileName = require("sanitize-filename");
import axios from 'axios';

import { SpreadSheet } from "./spreadsheet";

// The file
const outputDirectory = process.argv[2];

// OrganizationID
const organizationId = process.argv[3];

// You can get this by getting a link for a board: https://favro.com/widget/OrganizationID/WidgetCommonKey
const widgetCommonKey = process.argv[4];

// This is your email address (which you generated a token for)
let userEmail = process.argv[5];

// Generate this in your user profile
let apiToken = process.argv[6];

if (!outputDirectory || !widgetCommonKey || !apiToken || !organizationId || !userEmail)  {
	console.error("Missing parameters. Usage is:\n", "npm run execute outputDirectory organizationId widgetCommonKey userEmail apiToken");
	process.exit(1);
}

console.log("outputDirectory", outputDirectory);
console.log("organizationId", organizationId);
console.log("widgetCommonKey", widgetCommonKey);
console.log("userEmail", userEmail, "\n");

const baseURL = "https://favro.com/api/v1/";

const streamFinished = util.promisify(stream.finished);

export async function downloadFile(fileUrl: string, outputLocationPath: string): Promise<any> {
  const writer = fs.createWriteStream(outputLocationPath);
  return axios({
    method: 'get',
    url: fileUrl,
    responseType: 'stream',
  }).then(async response => {
    response.data.pipe(writer);
    return streamFinished(writer); //this is a Promise
  });
}

async function performRequest(method, subPath, postData?): Promise<any> {
	let options: https.RequestOptions = url.parse(baseURL + subPath);

	options.method = method;
	options.auth = userEmail + ":" + apiToken;
	options.headers = {
		organizationId,
		Accept: "application/json",
		"Content-Type": "application/json",
	};

	if (postData) {
		postData = JSON.stringify(postData);
		options.headers['Content-Length'] = Buffer.byteLength(postData);
	}

	let data = "";
	let statusCode = 0;
	let statusMessage = "";

	return new Promise(function(resolve, reject) {
		const req = https.request(options, (res) => {
			statusMessage = res.statusMessage;
			statusCode = res.statusCode;

			res.setEncoding("utf8");
			res.on("data", d => data += d);
		});

		req.on("error", function(e) {
			//console.error(e);
			reject(e);
		});

		req.on("close", () => {
			let resultJson: any = JSON.parse(data);
			if (statusCode >= 300) {
				if (resultJson.message)
					reject(resultJson.message);
				else
					reject(statusMessage);
			} else {
				resolve(resultJson);
			}
		});

		if (postData)
			req.write(postData);

		req.end();
	});
}

async function getAllCards(): Promise<any []> {
	let cards = [];

	let page = 0;
	let requestId: string;

	let startTime = moment();

	while (1) {
		let requestURL = `cards?widgetCommonId=${widgetCommonKey}&page=${page++}`;
		if (requestId)
			requestURL += "&requestId=" + requestId;

		let result: any = await performRequest("GET", requestURL);

		if (!requestId)
			requestId = result.requestId;

		for (let card of result.entities)
			cards.push(card);

		if (result.pages <= page)
			return cards;

		let now = moment();
		if (now.diff(startTime, "seconds") > 3) {
			startTime = now;
			console.log(`Getting cards ${(page / result.pages * 100).toFixed(0)}% done`);
		}
	}
}

async function getCustomField(fieldId: string): Promise<any> {
	let requestURL = `customfields/${fieldId}`;
	return await performRequest("GET", requestURL);
}

async function getWidget(widgetCommonKey: string): Promise<any> {
	let requestURL = `widgets/${widgetCommonKey}`;
	return await performRequest("GET", requestURL);
}

interface Column {
	id: string;
	name: string;
};

interface Attachment {
	cardId: string;
	cardName: string;
	name: string;
	url: string;
};

async function exportWidget(widgetCommonKey: string) {
	let allCards = await getAllCards();

	let fields: object = {};

	for (let card of allCards) {
		if (!card.customFields)
			continue;

		for (let field of card.customFields) {
			if (fields[field.customFieldId])
				continue;

			fields[field.customFieldId] = await getCustomField(field.customFieldId);
		}
	}

	let columns: Column[] = [];

	for (let fieldId in fields) {
		let field = fields[fieldId];

		columns.push({
			id: fieldId,
			name: field.name,
		});
	}

	let spreadsheet: SpreadSheet.Spreadsheet = {
		name: `Favro Export for widget ${widgetCommonKey} - ${moment().format("YYYY-MM-DD HH:MM")}`,
		columns: columns.reduce((array: string [], column: Column) => {
			array.push(column.name);
			return array;
		}, ["Name", "Detailed Description"]),
		data: [],
	};

	let attachmentsToDownload: Attachment[] = [];

	for (let card of allCards) {
		let cells: SpreadSheet.SpreadsheetCell[] = [
			{
				value: card.name,
				type: SpreadSheet.EExportReportFieldType.String,
			},
			{
				value: card.detailedDescription ? card.detailedDescription : "",
				type: SpreadSheet.EExportReportFieldType.String,
			},
		];

		if (card.attachments) {
			for (let attachment of card.attachments) {
				attachmentsToDownload.push({
					cardId: card.cardId,
					cardName: card.name,
					name: attachment.name,
					url: attachment.fileURL,
				});
			}
		}

		for (let column of columns) {
			let field = fields[column.id];
			let fieldValue: any;

			if (card.customFields) {
				for (let fieldData of card.customFields) {
					if (fieldData.customFieldId == column.id) {
						fieldValue = fieldData.value;
						break;
					}
				}
			}

			if (!fieldValue) {
				cells.push({
					value: "",
					type: SpreadSheet.EExportReportFieldType.String,
				});

				continue;
			}

			switch (field.type) {
				case "Number": {
					cells.push({
						value: fieldValue,
						type: SpreadSheet.EExportReportFieldType.Numeric,
					});
				}
				break;
				case "Text": {
					cells.push({
						value: fieldValue,
						type: SpreadSheet.EExportReportFieldType.String,
					});
				}
				break;
				case "Date": {
					cells.push({
						value: moment(fieldValue).toDate(),
						type: SpreadSheet.EExportReportFieldType.Date,
					});
				}
				break;
				case "Time": {
					cells.push({
						value: fieldValue.total,
						type: SpreadSheet.EExportReportFieldType.Numeric,
					});
				}
				break;
				case "Single select":
				case "Multiple select": {
					let itemValue: string [] = [];
					for (let value of fieldValue) {
						let found = false;
						for (let fieldItem of field.customFieldItems) {
							if (fieldItem.customFieldItemId == value) {
								itemValue.push(fieldItem.name);
								found = true;
								break;
							}
						}
						if (!found)
							itemValue.push("INVALID ITEM");
					}
					cells.push({
						value: itemValue.join(", "),
						type: SpreadSheet.EExportReportFieldType.String,
					});
				}
				break;
				default: {
					cells.push({
						value: `UNSUPPORTED - {field.type}`,
						type: SpreadSheet.EExportReportFieldType.String,
					});
				}
				break;
			}
		}

		spreadsheet.data.push({cells});
	}

	let widget = await getWidget(widgetCommonKey);
	let directory = `${outputDirectory}/${sanitizeFileName(`${widgetCommonKey} - ${widget.name}`)}`;

	await fs.promises.mkdir(directory, {recursive: true});
	SpreadSheet.writeToFile(spreadsheet, directory + "/sheet.xlsx");

	await fs.promises.writeFile(directory + "/allCards.json", JSON.stringify(allCards, null, 4));

	console.log(`Exported ${allCards.length} cards, now downloading ${attachmentsToDownload.length} attachments`);

	{
		let startTime = moment();

		let nFinished = 0;
		for (let attachment of attachmentsToDownload) {
			let attachmentDirectory = `${directory}/${sanitizeFileName(`${attachment.cardId} - ${attachment.cardName}`)}`
			await fs.promises.mkdir(attachmentDirectory, {recursive: true});

			await downloadFile(attachment.url, `${attachmentDirectory}/${sanitizeFileName(attachment.name)}`);
			++nFinished;

			let now = moment();
			if (now.diff(startTime, "seconds") > 3) {
				startTime = now;
				console.log(`Getting attachments ${(nFinished / attachmentsToDownload.length * 100).toFixed(2)}% done`);
			}
		}
	}
}

exportWidget(widgetCommonKey).then(function (){
	console.log("All done");
}).catch(function (error) {
	console.error("Failed", error.stack || error);
	process.exit(1);
});
