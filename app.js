import { randomBytes, scrypt as scryptCallback, timingSafeEqual } from 'node:crypto'
import { mkdirSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { createServer } from 'node:http'
import { basename, dirname, extname, join, resolve, sep } from 'node:path'
import { fileURLToPath } from 'node:url'
import { promisify } from 'node:util'
import { DatabaseSync } from 'node:sqlite'
import Busboy from 'busboy'
import { load as loadHtml } from 'cheerio'
import { Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, WidthType } from 'docx'
import ExcelJS from 'exceljs'
import { fileTypeFromBuffer } from 'file-type'
import mammoth from 'mammoth'
import { PDFDocument } from 'pdf-lib'
import { PDFParse } from 'pdf-parse'
import { createWorker } from 'tesseract.js'

const scrypt = promisify(scryptCallback)
const projectRoot = dirname(fileURLToPath(import.meta.url))
const dataDirectory = process.env.RAILWAY_VOLUME_MOUNT_PATH || join(projectRoot, 'data')
const databasePath = resolve(process.env.DB_PATH || join(dataDirectory, 'pdfflow.sqlite'))
const host = process.env.HOST || (process.env.PORT ? '0.0.0.0' : '127.0.0.1')
const port = Number(process.env.PORT || 5001)
const frontendRoot = resolve(projectRoot, 'dist')
const maxUploadBytes = 25 * 1024 * 1024
const maxRequestBytes = 50 * 1024 * 1024
const docxMimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const xlsxMimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
let ocrWorkerPromise

mkdirSync(dirname(databasePath), { recursive: true })

const database = new DatabaseSync(databasePath)
database.exec(`
	PRAGMA journal_mode = WAL;
	CREATE TABLE IF NOT EXISTS users (
		id INTEGER PRIMARY KEY,
		name TEXT NOT NULL,
		email TEXT NOT NULL UNIQUE COLLATE NOCASE,
		password_hash TEXT NOT NULL,
		password_salt TEXT NOT NULL,
		created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
	);
	CREATE TABLE IF NOT EXISTS page_views (
		id INTEGER PRIMARY KEY,
		type TEXT NOT NULL,
		path TEXT NOT NULL,
		created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
	);
`)

const findUserByEmail = database.prepare('SELECT * FROM users WHERE email = ?')
const createUser = database.prepare(
	'INSERT INTO users (name, email, password_hash, password_salt) VALUES (?, ?, ?, ?)',
)
const recordPageView = database.prepare('INSERT INTO page_views (type, path) VALUES (?, ?)')

function httpError(statusCode, message) {
	return Object.assign(new Error(message), { statusCode })
}

function sendJson(response, statusCode, body) {
	response.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' })
	response.end(JSON.stringify(body))
}

async function serveFrontend(pathname, response) {
	let decodedPath
	try {
		decodedPath = decodeURIComponent(pathname)
	} catch {
		throw httpError(400, 'Invalid URL path.')
	}

	const requestedPath = decodedPath === '/' ? '/index.html' : decodedPath
	const filePath = resolve(frontendRoot, `.${requestedPath}`)
	if (filePath !== frontendRoot && !filePath.startsWith(`${frontendRoot}${sep}`)) {
		throw httpError(404, 'Not found.')
	}

	let content
	let servedPath = filePath
	try {
		content = await readFile(filePath)
	} catch {
		if (extname(requestedPath)) throw httpError(404, 'Not found.')
		servedPath = join(frontendRoot, 'index.html')
		try {
			content = await readFile(servedPath)
		} catch {
			throw httpError(503, 'Frontend assets have not been built.')
		}
	}

	const contentTypes = {
		'.css': 'text/css; charset=utf-8',
		'.html': 'text/html; charset=utf-8',
		'.ico': 'image/x-icon',
		'.js': 'text/javascript; charset=utf-8',
		'.json': 'application/json; charset=utf-8',
		'.svg': 'image/svg+xml',
	}
	response.writeHead(200, {
		'Content-Type': contentTypes[extname(servedPath).toLowerCase()] || 'application/octet-stream',
		'Content-Length': content.length,
	})
	response.end(content)
}

async function readJson(request) {
	const chunks = []
	let size = 0

	for await (const chunk of request) {
		size += chunk.length
		if (size > 64 * 1024) {
			throw httpError(413, 'Request body is too large.')
		}
		chunks.push(chunk)
	}

	let body
	try {
		body = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}')
	} catch {
		throw httpError(400, 'Request body must be valid JSON.')
	}

	if (!body || typeof body !== 'object' || Array.isArray(body)) {
		throw httpError(400, 'Request body must be a JSON object.')
	}

	return body
}

function readMultipart(request) {
	return new Promise((resolveMultipart, rejectMultipart) => {
		let parser
		try {
			parser = Busboy({
				headers: request.headers,
				limits: { fileSize: maxUploadBytes, files: 5, fields: 8, fieldSize: 16 * 1024 },
			})
		} catch {
			rejectMultipart(httpError(400, 'Upload must use multipart form data.'))
			return
		}

		const files = []
		const fields = {}
		let totalBytes = 0
		let failure
		const fail = (error) => { failure ||= error }

		parser.on('field', (name, value, info) => {
			if (info.valueTruncated) {
				fail(httpError(413, 'An upload field is too large.'))
				return
			}
			fields[name] = value
		})
		parser.on('file', (field, stream, info) => {
			const chunks = []
			let fileBytes = 0
			stream.on('data', (chunk) => {
				fileBytes += chunk.length
				totalBytes += chunk.length
				if (fileBytes > maxUploadBytes || totalBytes > maxRequestBytes) {
					fail(httpError(413, 'The upload exceeds the file size limit.'))
				} else if (!failure) {
					chunks.push(chunk)
				}
			})
			stream.on('limit', () => fail(httpError(413, 'Each file must be smaller than 25 MB.')))
			stream.on('end', () => {
				if (!stream.truncated && !failure && info.filename) {
					files.push({
						field,
						filename: info.filename,
						mimeType: info.mimeType,
						buffer: Buffer.concat(chunks),
					})
				}
			})
		})
		parser.on('filesLimit', () => fail(httpError(413, 'Upload no more than 5 files at once.')))
		parser.on('fieldsLimit', () => fail(httpError(413, 'Too many upload fields.')))
		parser.on('error', rejectMultipart)
		parser.on('close', () => {
			if (failure) {
				rejectMultipart(failure)
			} else {
				resolveMultipart({ files, fields })
			}
		})
		request.pipe(parser)
	})
}

function normalizeEmail(value) {
	return typeof value === 'string' ? value.trim().toLowerCase() : ''
}

function safeFileBaseName(filename) {
	const name = basename(filename).split(/[\\/]/).pop().replace(extname(filename), '')
	return name.replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 100) || 'converted'
}

function sendFile(response, buffer, filename, mimeType) {
	const safeName = filename.replace(/[\r\n"]/g, '_')
	response.writeHead(200, {
		'Content-Type': mimeType,
		'Content-Length': buffer.length,
		'Content-Disposition': `attachment; filename="${safeName}"`,
	})
	response.end(buffer)
}

function requireFile(file, field = 'file') {
	if (!file || !file.buffer.length) {
		throw httpError(400, `Choose a ${field} to convert.`)
	}
	return extname(file.filename).toLowerCase()
}

async function requirePdf(file) {
	if (requireFile(file) !== '.pdf') {
		throw httpError(415, 'Choose a PDF file.')
	}
	const detected = await fileTypeFromBuffer(file.buffer)
	if (detected?.mime !== 'application/pdf') {
		throw httpError(415, 'The uploaded file is not a valid PDF.')
	}
}

function requireOfficeZip(file, extension, label) {
	if (requireFile(file) !== extension || !file.buffer.subarray(0, 4).equals(Buffer.from([0x50, 0x4b, 0x03, 0x04]))) {
		throw httpError(415, `Choose a valid ${label} file (${extension}).`)
	}
}

function paragraphsFromText(text) {
	return text.split(/\r?\n/).map((line) => new Paragraph({ text: line || ' ' }))
}

async function convertPdfToWord(file) {
	await requirePdf(file)
	const parser = new PDFParse({ data: file.buffer })
	try {
		const result = await parser.getText()
		if (!result.text.trim()) {
			throw httpError(422, 'This PDF has no selectable text. Scanned PDFs need OCR before conversion.')
		}
		const document = new Document({ sections: [{ children: paragraphsFromText(result.text.trim()) }] })
		return Packer.toBuffer(document)
	} catch (error) {
		if (error.statusCode) throw error
		throw httpError(422, 'Could not extract text from this PDF. It may be encrypted or damaged.')
	} finally {
		await parser.destroy()
	}
}

function rowsFromWordHtml(html) {
	const $ = loadHtml(`<div id="document-root">${html}</div>`)
	const root = $('#document-root')
	const rows = []
	for (const node of root.contents().toArray()) {
		if (node.type !== 'tag') continue
		if (node.name === 'table') {
			for (const tableRow of $(node).find('tr').toArray()) {
				const cells = $(tableRow).find('th, td').toArray().map((cell) => $(cell).text().trim())
				if (cells.length) rows.push(cells)
			}
		} else {
			const text = $(node).text().trim()
			if (text) rows.push([text])
		}
	}
	return rows
}

async function convertWordToExcel(file) {
	requireOfficeZip(file, '.docx', 'Word')
	try {
		const result = await mammoth.convertToHtml({ buffer: file.buffer })
		const rows = rowsFromWordHtml(result.value)
		if (!rows.length) throw httpError(422, 'This Word document contains no text or tables to export.')
		const workbook = new ExcelJS.Workbook()
		const worksheet = workbook.addWorksheet('Converted')
		rows.forEach((row) => worksheet.addRow(row))
		return Buffer.from(await workbook.xlsx.writeBuffer())
	} catch (error) {
		if (error.statusCode) throw error
		throw httpError(422, 'Could not read this Word document. Use a valid .docx file.')
	}
}

function excelCellText(value) {
	if (value == null) return ''
	if (value instanceof Date) return value.toISOString()
	if (typeof value === 'object') {
		if ('text' in value) return String(value.text)
		if ('richText' in value) return value.richText.map((part) => part.text).join('')
		if ('result' in value) return excelCellText(value.result)
		return JSON.stringify(value)
	}
	return String(value)
}

async function convertExcelToWord(file) {
	requireOfficeZip(file, '.xlsx', 'Excel')
	try {
		const workbook = new ExcelJS.Workbook()
		await workbook.xlsx.load(file.buffer)
		const children = []
		for (const worksheet of workbook.worksheets) {
			const rows = []
			worksheet.eachRow({ includeEmpty: false }, (row) => {
				rows.push(row.values.slice(1).map(excelCellText))
			})
			if (!rows.length) continue
			children.push(new Paragraph({ text: worksheet.name, heading: HeadingLevel.HEADING_1 }))
			const columnCount = Math.max(...rows.map((row) => row.length))
			const tableRows = rows.map((row) => new TableRow({
				children: Array.from({ length: columnCount }, (_, index) => new TableCell({
					children: [new Paragraph({ text: row[index] || ' ' })],
				})),
			}))
			children.push(new Table({
				rows: tableRows,
				width: { size: 100, type: WidthType.PERCENTAGE },
			}))
		}
		if (!children.length) throw httpError(422, 'This spreadsheet contains no data to export.')
		return Packer.toBuffer(new Document({ sections: [{ children }] }))
	} catch (error) {
		if (error.statusCode) throw error
		throw httpError(422, 'Could not read this spreadsheet. Use a valid .xlsx file.')
	}
}

async function recognizeImage(file) {
	const extension = requireFile(file)
	const expectedMime = extension === '.png' ? 'image/png' : ['.jpg', '.jpeg'].includes(extension) ? 'image/jpeg' : ''
	const detected = await fileTypeFromBuffer(file.buffer)
	if (!expectedMime || detected?.mime !== expectedMime) {
		throw httpError(415, 'Choose a valid PNG or JPEG image.')
	}
	try {
		ocrWorkerPromise ||= createWorker('eng')
		const worker = await ocrWorkerPromise
		const result = await worker.recognize(file.buffer)
		if (!result.data.text.trim()) throw httpError(422, 'No readable text was found in this image.')
		return result.data.text.trim()
	} catch (error) {
		if (error.statusCode) throw error
		ocrWorkerPromise = undefined
		throw httpError(503, 'The OCR engine could not process this image. Check the server network and try again.')
	}
}

async function convertImageToWord(file) {
	const text = await recognizeImage(file)
	const document = new Document({ sections: [{ children: paragraphsFromText(text) }] })
	return Packer.toBuffer(document)
}

function parsePageSelection(value, pageCount) {
	if (typeof value !== 'string' || !value.trim()) {
		throw httpError(400, 'Enter page numbers or ranges, such as 1-3,5.')
	}
	const pages = []
	for (const item of value.split(',')) {
		const match = /^(\d+)(?:\s*-\s*(\d+))?$/.exec(item.trim())
		if (!match) throw httpError(400, 'Use page numbers or ranges such as 1-3,5.')
		const first = Number(match[1])
		const last = Number(match[2] || match[1])
		if (first < 1 || last < first || last > pageCount) {
			throw httpError(400, `Page numbers must be between 1 and ${pageCount}.`)
		}
		for (let page = first; page <= last; page += 1) pages.push(page - 1)
	}
	return [...new Set(pages)]
}

async function loadPdfDocument(file) {
	await requirePdf(file)
	try {
		return await PDFDocument.load(file.buffer)
	} catch {
		throw httpError(422, `Could not read ${basename(file.filename)}. The PDF may be encrypted or damaged.`)
	}
}

async function handlePdfTool(action, files, fields, response) {
	if (action === 'merge') {
		if (files.length < 2) throw httpError(400, 'Choose at least two PDF files to merge.')
		const output = await PDFDocument.create()
		for (const file of files) {
			const source = await loadPdfDocument(file)
			const pages = await output.copyPages(source, source.getPageIndices())
			pages.forEach((page) => output.addPage(page))
		}
		sendFile(response, Buffer.from(await output.save()), 'merged.pdf', 'application/pdf')
		return
	}
	if (action === 'extract') {
		if (files.length !== 1) throw httpError(400, 'Choose one PDF file to extract pages from.')
		const source = await loadPdfDocument(files[0])
		const selectedPages = parsePageSelection(fields.pages, source.getPageCount())
		const output = await PDFDocument.create()
		const pages = await output.copyPages(source, selectedPages)
		pages.forEach((page) => output.addPage(page))
		const filename = `${safeFileBaseName(files[0].filename)}-pages.pdf`
		sendFile(response, Buffer.from(await output.save()), filename, 'application/pdf')
		return
	}
	throw httpError(404, 'Unknown PDF tool.')
}

async function handleConversion(tool, files, response) {
	if (files.length !== 1) throw httpError(400, 'Choose one file to convert.')
	const file = files[0]
	let output
	let filename
	let mimeType
	if (tool === 'pdf-to-word') {
		output = await convertPdfToWord(file)
		filename = `${safeFileBaseName(file.filename)}.docx`
		mimeType = docxMimeType
	} else if (tool === 'word-to-excel') {
		output = await convertWordToExcel(file)
		filename = `${safeFileBaseName(file.filename)}.xlsx`
		mimeType = xlsxMimeType
	} else if (tool === 'excel-to-word') {
		output = await convertExcelToWord(file)
		filename = `${safeFileBaseName(file.filename)}.docx`
		mimeType = docxMimeType
	} else if (tool === 'image-to-word') {
		output = await convertImageToWord(file)
		filename = `${safeFileBaseName(file.filename)}.docx`
		mimeType = docxMimeType
	} else {
		throw httpError(404, 'Unknown conversion tool.')
	}
	sendFile(response, output, filename, mimeType)
}

async function handleRequest(request, response) {
	const origin = request.headers.origin
	if (origin && /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)) {
		response.setHeader('Access-Control-Allow-Origin', origin)
		response.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
		response.setHeader('Access-Control-Allow-Headers', 'Content-Type')
		response.setHeader('Access-Control-Expose-Headers', 'Content-Disposition')
		response.setHeader('Vary', 'Origin')
	}

	if (request.method === 'OPTIONS') {
		response.writeHead(204)
		response.end()
		return
	}

	const pathname = new URL(request.url, 'http://localhost').pathname

	try {
		if (request.method === 'GET') {
			if (pathname === '/health') {
				sendJson(response, 200, { status: 'ok' })
				return
			}
			if (pathname.startsWith('/api/')) {
				sendJson(response, 404, { message: 'Not found.' })
				return
			}
			await serveFrontend(pathname, response)
			return
		}
		if (request.method !== 'POST') {
			sendJson(response, 404, { message: 'Not found.' })
			return
		}
		if (pathname.startsWith('/api/convert/')) {
			const { files } = await readMultipart(request)
			await handleConversion(pathname.slice('/api/convert/'.length), files, response)
			return
		}
		if (pathname.startsWith('/api/pdf-tools/')) {
			const { files, fields } = await readMultipart(request)
			await handlePdfTool(pathname.slice('/api/pdf-tools/'.length), files, fields, response)
			return
		}

		const body = await readJson(request)

		if (pathname === '/api/traffic/record') {
			if (body.type !== 'pageview' || typeof body.path !== 'string' || body.path.length > 2048) {
				throw httpError(400, 'Invalid page-view data.')
			}

			recordPageView.run(body.type, body.path)
			sendJson(response, 201, { message: 'Page view recorded.' })
			return
		}

		if (pathname === '/api/signup') {
			const name = typeof body.name === 'string' ? body.name.trim() : ''
			const email = normalizeEmail(body.email)
			const password = typeof body.password === 'string' ? body.password : ''

			if (!name || name.length > 100 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
				throw httpError(400, 'Enter a valid name and email address.')
			}
			if (password.length < 8 || password.length > 128) {
				throw httpError(400, 'Password must be between 8 and 128 characters.')
			}
			if (findUserByEmail.get(email)) {
				throw httpError(409, 'An account with this email already exists.')
			}

			const salt = randomBytes(16)
			const passwordHash = await scrypt(password, salt, 64)
			createUser.run(name, email, passwordHash.toString('hex'), salt.toString('hex'))
			sendJson(response, 201, { message: 'Account created successfully.' })
			return
		}

		if (pathname === '/api/signin') {
			const email = normalizeEmail(body.email)
			const password = typeof body.password === 'string' ? body.password : ''

			if (!email || !password || password.length > 128) {
				throw httpError(401, 'Invalid email or password.')
			}

			const user = findUserByEmail.get(email)
			const salt = user ? Buffer.from(user.password_salt, 'hex') : Buffer.alloc(16)
			const expectedHash = user ? Buffer.from(user.password_hash, 'hex') : Buffer.alloc(64)
			const actualHash = await scrypt(password, salt, 64)

			if (!user || !timingSafeEqual(expectedHash, actualHash)) {
				throw httpError(401, 'Invalid email or password.')
			}

			sendJson(response, 200, { message: 'Signed in successfully.' })
			return
		}

		sendJson(response, 404, { message: 'Not found.' })
	} catch (error) {
		if (!error.statusCode) {
			console.error('Request failed:', error)
		}
		sendJson(response, error.statusCode || 500, {
			message: error.statusCode ? error.message : 'The server could not process the request.',
		})
	}
}

const server = createServer(handleRequest)
server.listen(port, host, () => {
	console.log(`PDFFlow API listening at http://${host}:${port}`)
	console.log(`User data is stored in ${databasePath}`)
})

for (const signal of ['SIGINT', 'SIGTERM']) {
	process.on(signal, () => {
		server.close(() => {
				Promise.resolve(ocrWorkerPromise)
					.then((worker) => worker?.terminate())
					.catch(() => {})
					.finally(() => {
						database.close()
						process.exit(0)
					})
		})
	})
}
