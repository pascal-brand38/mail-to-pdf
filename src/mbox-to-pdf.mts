#!/usr/bin/env node

// Copyright (c) Pascal Brand
// MIT License

import * as fs from 'fs'
import * as path from 'path'

// @ts-ignore
import { mboxReader } from 'mbox-reader'  // scan messages

import { simpleParser, ParsedMail, HeaderValue, AddressObject } from 'mailparser' // parse a single message
import puppeteer from "puppeteer"       // save html text as pdf

const mboxPath = 'C:/tmp/mbox-to-pdf/INBOX'
const outputDir = 'C:/tmp/mbox-to-pdf/output'


const readStream = fs.createReadStream(mboxPath)

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

  return header
}


function getHtml(parser: ParsedMail, header: Header): string {
  let bodyStr = ''
  if (parser.html) {
    bodyStr = parser.html
  }

  let html = ''
  html += `<div style="background-color:lightgrey;">`
  html += '<div><br></div>'
  html += '<div><em>Generated using https://npmjs.com/package/mbox-to-pdf</em></div>'
  html += '<div><br></div>'
  html += `<div><strong>From:</strong> ${escape(header.from)}</div>`
  html += `<div><strong>To:</strong> ${escape(header.to)}</div>`
  html += `<div><strong>Cc:</strong> ${escape(header.cc)}</div>`
  html += `<div><strong>Bcc:</strong> ${escape(header.bcc)}</div>`
  html += `<div><strong>Subject:</strong> ${escape(header.subject)}</div>`
  html += `<div><strong>Date:</strong> ${escape(header.date)}</div>`
  html += '<div><br></div>'
  html += `</div>`
  html += '<div><br></div>'
  html += bodyStr
  return html
}


const browser = await puppeteer.launch();

for await (let message of mboxReader(readStream)) {
  const parser = await simpleParser(message.content);
  const header = getHeader(parser)

  const targetDir = path.join(outputDir, header.basename)
  fs.mkdirSync(targetDir, { recursive: true });

  parser.attachments.forEach((attachment, index) => {
    if (attachment.filename) {
      fs.writeFileSync(path.join(targetDir, attachment.filename), attachment.content);
    } else {
      fs.writeFileSync(path.join(targetDir, `/attach-${index}`), attachment.content);
      console.log('ERROR:', header)
    }
  })

  if (parser.html) {
    const page = await browser.newPage();
    //set the HTML of this page
    await page.setContent(getHtml(parser, header));
    //save the page into a PDF and call it 'puppeteer-example.pdf'
    await page.pdf({ path: path.join(targetDir, header.basename+'.pdf'), printBackground: true });
  }
}

await browser.close();
console.log('DONE')


// TODO
// - eml is attached => no filename of the attachment, and the real attachment is in the eml that is attached
// - add link to attachment in the body.pdf
