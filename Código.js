function doGet() {
	setupContrateBackend();
	setupContrateUsersBackend();
	return HtmlService.createHtmlOutputFromFile('Index')
		.setTitle('CONTRATE | Acceso corporativo')
		.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const BACKEND_CONFIG = {
	propertyKey: 'CONTRATE_SPREADSHEET_ID',
	sheetName: 'Registros',
	spreadsheetName: 'CONTRATE | Base de datos solicitudes_marketing'
};

const RECORD_HEADERS = [
	'id',
	'psicologo',
	'cargo',
	'ciudad',
	'funciones',
	'requisitos',
	'habilidades',
	'salario',
	'contrato',
	'beneficios',
	'horarios',
	'correo',
	'wtpp',
	'status',
	'createdAt',
	'updatedAt'
];

const USER_HEADERS = [
	'email',
	'password',
	'role',
	'name',
	'active'
];

const DEFAULT_USERS = [
	{
		email: 'psicologia@contrate.com',
		password: 'Contrate2026',
		role: 'psicologia',
		name: 'Psicologia Demo',
		active: 'TRUE'
	},
	{
		email: 'marketing@contrate.com',
		password: 'Contrate2026',
		role: 'marketing',
		name: 'Marketing Demo',
		active: 'TRUE'
	}
];

const USERS_BACKEND_CONFIG = {
	propertyKey: 'CONTRATE_SPREADSHEET_ID',
	sheetName: 'Usuarios'
};

function setupContrateBackend() {
	const spreadsheet = ensureSpreadsheet_();
	return {
		spreadsheetId: spreadsheet.getId(),
		spreadsheetUrl: spreadsheet.getUrl(),
		sheetName: BACKEND_CONFIG.sheetName
	};
}

function getRecordsFromSheet() {
	const sheet = getRecordsSheet_();
	const values = sheet.getDataRange().getValues();
	if (values.length <= 1) {
		return [];
	}

	const headers = values[0].map((header) => String(header || '').trim());
	return values.slice(1)
		.filter((row) => row.some((cell) => cell !== '' && cell !== null))
		.map((row) => rowToRecord_(headers, row));
}

function saveRecordsToSheet(records) {
	const sheet = getRecordsSheet_();
	sheet.clearContents();
	sheet.getRange(1, 1, 1, RECORD_HEADERS.length).setValues([RECORD_HEADERS]);

	const rows = Array.isArray(records)
		? records.map((record) => RECORD_HEADERS.map((header) => normalizeCellValue_(record ? record[header] : '')))
		: [];

	if (rows.length) {
		sheet.getRange(2, 1, rows.length, RECORD_HEADERS.length).setValues(rows);
	}

	return { saved: rows.length };
}

function ensureSpreadsheet_() {
	const properties = PropertiesService.getScriptProperties();
	const storedId = properties.getProperty(BACKEND_CONFIG.propertyKey);

	if (storedId) {
		try {
			const spreadsheet = SpreadsheetApp.openById(storedId);
			ensureSheetHeader_(spreadsheet);
			return spreadsheet;
		} catch (error) {
			properties.deleteProperty(BACKEND_CONFIG.propertyKey);
		}
	}

	const spreadsheet = SpreadsheetApp.create(BACKEND_CONFIG.spreadsheetName);
	const sheet = spreadsheet.getSheets()[0];
	sheet.setName(BACKEND_CONFIG.sheetName);
	sheet.clearContents();
	sheet.getRange(1, 1, 1, RECORD_HEADERS.length).setValues([RECORD_HEADERS]);
	properties.setProperty(BACKEND_CONFIG.propertyKey, spreadsheet.getId());
	return spreadsheet;
}

function getRecordsSheet_() {
	const spreadsheet = ensureSpreadsheet_();
	let sheet = spreadsheet.getSheetByName(BACKEND_CONFIG.sheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(BACKEND_CONFIG.sheetName);
	}
	ensureSheetHeader_(spreadsheet);
	return sheet;
}

function ensureSheetHeader_(spreadsheet) {
	const sheet = spreadsheet.getSheetByName(BACKEND_CONFIG.sheetName) || spreadsheet.getSheets()[0];
	const headerRange = sheet.getRange(1, 1, 1, RECORD_HEADERS.length);
	const currentHeaders = headerRange.getValues()[0].map((header) => String(header || '').trim());
	const headerMatches = RECORD_HEADERS.every((header, index) => currentHeaders[index] === header);

	if (!headerMatches) {
		sheet.getRange(1, 1, 1, RECORD_HEADERS.length).setValues([RECORD_HEADERS]);
	}
}

function rowToRecord_(headers, row) {
	return headers.reduce((record, header, index) => {
		if (header) {
			record[header] = row[index];
		}
		return record;
	}, {});
}

function normalizeCellValue_(value) {
	if (value === null || value === undefined) {
		return '';
	}

	if (value instanceof Date) {
		return value.toISOString();
	}

	if (typeof value === 'object') {
		return JSON.stringify(value);
	}

	return String(value);
}

function include(filename) {
	return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function myFunction() {
	return doGet();
}

function setupContrateUsersBackend() {
	const spreadsheet = ensureSpreadsheet_();
	const sheet = getUsersSheet_(spreadsheet);
	seedUsersSheet_(sheet);
	return {
		spreadsheetId: spreadsheet.getId(),
		spreadsheetUrl: spreadsheet.getUrl(),
		sheetName: USERS_BACKEND_CONFIG.sheetName
	};
}

function getUsersFromSheet() {
	const sheet = getUsersSheet_();
	const values = sheet.getDataRange().getValues();
	if (values.length <= 1) {
		return [];
	}

	const headers = values[0].map((header) => String(header || '').trim());
	return values.slice(1)
		.filter((row) => row.some((cell) => cell !== '' && cell !== null))
		.map((row) => rowToRecord_(headers, row));
}

function saveUsersToSheet(users) {
	const sheet = getUsersSheet_();
	sheet.clearContents();
	sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);

	const rows = Array.isArray(users)
		? users.map((user) => USER_HEADERS.map((header) => normalizeCellValue_(user ? user[header] : '')))
		: [];

	if (rows.length) {
		sheet.getRange(2, 1, rows.length, USER_HEADERS.length).setValues(rows);
	}

	return { saved: rows.length };
}

function ensureUsersSpreadsheet_() {
	return ensureSpreadsheet_();
}

function getUsersSheet_(spreadsheet) {
	const targetSpreadsheet = spreadsheet || ensureUsersSpreadsheet_();
	let sheet = targetSpreadsheet.getSheetByName(USERS_BACKEND_CONFIG.sheetName);
	if (!sheet) {
		sheet = targetSpreadsheet.insertSheet(USERS_BACKEND_CONFIG.sheetName);
	}
	ensureUsersSheetHeader_(targetSpreadsheet);
	return sheet;
}

function ensureUsersSheetHeader_(spreadsheet) {
	const sheet = spreadsheet.getSheetByName(USERS_BACKEND_CONFIG.sheetName) || spreadsheet.getSheets()[0];
	const headerRange = sheet.getRange(1, 1, 1, USER_HEADERS.length);
	const currentHeaders = headerRange.getValues()[0].map((header) => String(header || '').trim());
	const headerMatches = USER_HEADERS.every((header, index) => currentHeaders[index] === header);

	if (!headerMatches) {
		sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
	}
}

function seedUsersSheet_(sheet) {
	const values = sheet.getDataRange().getValues();
	const hasUsers = values.length > 1;
	if (hasUsers) {
		return;
	}

	const rows = DEFAULT_USERS.map((user) => USER_HEADERS.map((header) => normalizeCellValue_(user[header])));
	sheet.clearContents();
	sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
	sheet.getRange(2, 1, rows.length, USER_HEADERS.length).setValues(rows);
}
