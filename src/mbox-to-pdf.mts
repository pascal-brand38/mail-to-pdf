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
  header.basename = header.basename.trim()

  return header
}


async function saveUsingPuppeteer(parser: ParsedMail, header: Header, targetDir: string, browser: Browser) {
  const pdfFullName = path.join(targetDir, header.basename+'.pdf')

  const page = await browser.newPage();
  await page.setContent(getHtml(parser, header));   //set the HTML of this page
  await page.pdf({ path: pdfFullName, printBackground: true });   // save the page

  let pdf = await PDFDocument.load(fs.readFileSync(pdfFullName))
  const pdfBuf = await pdf.save()
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


async function mboxToPdf(mboxPath: string, outputDir: string) {
  const readStream = fs.createReadStream(mboxPath)

  const browser = await puppeteer.launch();

  for await (let message of mboxReader(readStream)) {
    const parser = await simpleParser(message.content);
    const header = getHeader(parser)

    // console.log(parser)

    console.log(`--- ${header.basename}`)

    const targetDir = path.join(outputDir, header.basename)
    fs.mkdirSync(targetDir, { recursive: true });

    parser.attachments.forEach((attachment, index) => {
      let filename = attachment.filename
      if (!filename &&  (attachment.contentType === 'message/rfc822')) {
        console.log('EML as attachment: ', header)
        filename = `/attachment-${index}.eml`
      }

      if (!filename) {
        console.log('ERROR attachment without filename: ', header)
        filename = `/attachment-${index}.unknown`
      }

      fs.writeFileSync(path.join(targetDir, filename), attachment.content)
    })

    await saveUsingPuppeteer(parser, header, targetDir, browser)
  }

  await browser.close();
}


await mboxToPdf('C:/tmp/mbox-to-pdf/INBOX', 'C:/tmp/mbox-to-pdf/output')

console.log('DONE')


// TODO
// - eml is attached => no filename of the attachment, and the real attachment is in the eml that is attached
