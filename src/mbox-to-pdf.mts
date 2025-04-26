#!/usr/bin/env node

// Copyright (c) Pascal Brand
// MIT License

import * as fs from 'fs'
import * as path from 'path'

// @ts-ignore
import { mboxReader } from 'mbox-reader'  // scan messages

import { simpleParser, ParsedMail, AddressObject } from 'mailparser' // parse a single message
import puppeteer, { Browser } from "puppeteer"       // save html text as pdf
import { PDFDocument } from 'pdf-lib'
import { exit } from 'process'


function escape(s: string): string {
  return s.replace(
      /[^0-9A-Za-z ]/g,
      c => "&#" + c.charCodeAt(0) + ";"
  );
}

interface Header  {
  from: string,
  to: string,
  cc: string,
  bcc: string,
  subject: string,
  date: string,
  basename: string,
}

function fixFilename(filename: string) {
  return  filename.replace(/[\:\\\/\*\?\"\<\>\|]/g, '').trim()
}

function getHeader(parser: ParsedMail): Header {
  const header = {
    from: '',
    to: '',
    cc: '',
    bcc: '',
    subject: '',
    date: '',
    basename: ''
  }

  const from = parser.headers.get('from')
  if (from) {
    header.from = (from as AddressObject).text
  }
  const to = parser.headers.get('to')
  if (to) {
    header.to = (to as AddressObject).text
  }
  const cc = parser.headers.get('cc')
  if (cc) {
    header.cc = (cc as AddressObject).text
  }
  const bcc = parser.headers.get('bcc')
  if (bcc) {
    header.bcc = (bcc as AddressObject).text
  }
  const date = parser.headers.get('date')
  if (date) {
    const d: Date = (date as Date)
    header.date = d.toLocaleString()
    header.basename = `${d.getFullYear()}-${(d.getMonth()+1).toLocaleString(undefined, {minimumIntegerDigits: 2})}-${d.getDate().toLocaleString(undefined, {minimumIntegerDigits: 2})}`
    header.basename += `-${d.getHours().toLocaleString(undefined, {minimumIntegerDigits: 2})}.${d.getMinutes().toLocaleString(undefined, {minimumIntegerDigits: 2})}.${d.getSeconds().toLocaleString(undefined, {minimumIntegerDigits: 2})}`
  }

  const subject = parser.headers.get('subject')
  if (subject) {
    header.subject = (subject as string)
    header.basename += ` - ${header.subject.replace(/[\:\\\/\*\?\"\<\>\|]/g, '')}`
  }
  header.basename = fixFilename(header.basename)

  return header
}


async function saveUsingPuppeteer(parser: ParsedMail, header: Header, targetDir: string, browser: Browser) {
  const pdfFullName = path.join(targetDir, header.basename+'.pdf')

  const page = await browser.newPage();
  await page.setContent(getHtml(parser, header));   //set the HTML of this page
  await page.pdf({ path: pdfFullName, printBackground: true });   // save the page

  let pdf = await PDFDocument.load(fs.readFileSync(pdfFullName))
  const pdfBuf = await pdf.save()
  await page.close()
  fs.writeFileSync(pdfFullName, pdfBuf);
}

function beautifulSize(s: number) {
  if (s < 1024) {
    return `${s.toFixed(2)}Bytes`
  } else if (s < 1024*1024) {
    return `${(s/1024).toFixed(2)}KB`
  } else if (s < 1024*1024*1024) {
    return `${(s/(1024*1024)).toFixed(2)}MB`
  } else {
    return `${(s/(1024*1024*1024)).toFixed(2)}GB`
  }
}

function getHtml(parser: ParsedMail, header: Header): string {
  let bodyStr = ''
  if (parser.html) {
    bodyStr = parser.html
  } else if (parser.text) {
    bodyStr = `<div>${parser.text.replaceAll('\n', '<br>')}</div>`
  } else if (parser.textAsHtml) {
    bodyStr = parser.textAsHtml
    console.log('ERROR - textAsHtml:', header)
  }

  let html = ''
  html += `<div style="background-color:lightgrey;">`
  html += '<div><br></div>'
  html += '<div><em>Generated using https://npmjs.com/package/mbox-to-pdf</em></div>'
  html += '<div><br></div>'
  html += `<div>From: ${escape(header.from)}</div>`
  html += `<div>To: ${escape(header.to)}</div>`
  html += `<div>Cc: ${escape(header.cc)}</div>`
  html += `<div>Bcc: ${escape(header.bcc)}</div>`
  html += `<div>Subject: ${escape(header.subject)}</div>`
  html += `<div>Date: ${escape(header.date)}</div>`
  html += '<div><br></div>'
  html += `</div>`
  html += '<div><br></div>'

  html += bodyStr
  html += '<div><br></div>'

  if (parser.attachments.length !== 0) {
    html += `<div style="background-color:lightgrey;">`
    html += '<div><br></div>'
    parser.attachments.forEach((attachment, index) => {
      html += `<div>`
      if (attachment.filename) {
        html += `attachments: ${attachment.filename}`
      } else {
        html += `attachments: unknown`
      }
      html += ` ${beautifulSize(attachment.content.length)}`
      html += `</div>`
    })
    html += '<div><br></div>'
    html += `</div>`
  }
  return html
}

function filenameFromContentType(contentType: string, index: number, header: Header) {
  let extension = undefined
  if (contentType === 'message/rfc822') {
    extension = 'eml'
  } else if (contentType === 'application/pdf') {
    extension = 'pdf'
  } else if (contentType === 'image/jpeg') {
    extension = 'jpg'
  } else if (contentType === 'image/png') {
    extension = 'png'
  } else if (contentType === 'image/gif') {
    extension = 'gif'
  } else if (contentType === 'text/calendar') {
    extension = 'ics'
  } else {
    extension = 'unknown'
    console.log('ERROR attachment without filename: ', header)
    console.log(`attachment.contentType = ${contentType}`)
  }

  return `attachment-${index}.${extension}`
}

async function mboxToPdf(mboxPath: string, outputDir: string) {
  let lenWhite = 3
  console.log(`Processing mbox file: ${mboxPath}`)
  console.log(`Creating outputs in: ${outputDir}`)

  const readStream = fs.createReadStream(mboxPath)

  const browser = await puppeteer.launch();

  for await (let message of mboxReader(readStream)) {
    const parser = await simpleParser(message.content);
    const header = getHeader(parser)

    // console.log(parser)

    // console.log(`--- ${header.basename}`)
    process.stdout.write(`${" ".repeat(lenWhite)}\r`)
    lenWhite = 10 + header.basename.length
    process.stdout.write(`--- ${header.basename}\r`)


    const targetDir = path.join(outputDir, header.basename)

    // check if it already exists. If so, do not regenerate anything
    // TODO: --force option
    const pdfFullName = path.join(targetDir, header.basename+'.pdf')
    if (fs.existsSync(pdfFullName)) {
      continue
    }

    fs.mkdirSync(targetDir, { recursive: true });

    parser.attachments.forEach((attachment, index) => {
      let filename = attachment.filename
      if (!filename) {
        filename = filenameFromContentType(attachment.contentType, index, header)
      }
      filename = fixFilename(filename)
      fs.writeFileSync(path.join(targetDir, filename), attachment.content)
    })

    await saveUsingPuppeteer(parser, header, targetDir, browser)
  }

  await browser.close();
}

function getDirectories(source: string) {
  try {
    return fs.readdirSync(source, { withFileTypes: true })
      .filter(dirent => dirent.isDirectory())
      .map(dirent => dirent.name)
  } catch {
    return []
  }
}

interface mboxDesc { name: string, subdir: string}
function getMboxPaths(root: string, subdir: string = '.') {
  let results: mboxDesc[] = []
  const dir = path.join(root, subdir)
  try {
    const contents = fs.readdirSync(dir, { withFileTypes: true })
    contents.forEach(c => {
      if (c.isDirectory()) {
        results = [ ...results, ...getMboxPaths(root, path.join(subdir, c.name))]
      } else if (c.isFile()) {
        if (!c.name.endsWith('.msf')) {
          results.push({name: c.name, subdir: subdir})
        }
      }
    })
  } catch {
    return []
  }
  return results
}


if (!process.env.APPDATA) {
  throw '$APPDATA is not defined in the environment variables'
}
const thunderbirdProfileDir = path.join(process.env.APPDATA, 'Thunderbird', 'Profiles')
const allThunderbirdProfiles = getDirectories(thunderbirdProfileDir)

console.log(allThunderbirdProfiles)

allThunderbirdProfiles.map(async (profileDir) =>{
  const imapMailPath = path.join(thunderbirdProfileDir, profileDir, 'ImapMail')
  // console.log(imapMailPath, getDirectories(imapMailPath))
  // console.log(imapMailPath, getMboxPaths(imapMailPath))

  // getMboxPaths(imapMailPath).map(async (desc) => {
  //   // console.log(
  //   //   path.join('C:/tmp/mbox-to-pdf/output', desc.subdir),
  //   //   path.join(imapMailPath, desc.subdir, desc.name))
  //   await mboxToPdf(
  //     path.join('C:/tmp/mbox-to-pdf/output', desc.subdir),
  //     path.join(imapMailPath, desc.subdir, desc.name))
  // })

  for await (let desc of getMboxPaths(imapMailPath)) {
    // await new Promise(r => setTimeout(r, 10 * 1000)),
    // console.log(
    //   path.join('C:/tmp/mbox-to-pdf/output', desc.subdir),
    //   path.join(imapMailPath, desc.subdir, desc.name))
    await mboxToPdf(
      path.join(imapMailPath, desc.subdir, desc.name),
      path.join('C:/tmp/mbox-to-pdf/output', desc.subdir, desc.name))
  }
})


// await mboxToPdf('C:/tmp/mbox-to-pdf/INBOX', 'C:/tmp/mbox-to-pdf/output')

console.log('DONE')


// TODO
// - remove 'Important' which are not duplicate (in case they do not have labels)
// - remove 'Tous les messages' which are not duplicate (in case they do not have labels)
// - 'Tous les message' may have another name in other language
