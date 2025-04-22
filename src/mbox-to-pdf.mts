#!/usr/bin/env node

// Copyright (c) Pascal Brand
// MIT License

import * as fs from 'fs'

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

function getMailBaseName(parser: ParsedMail): string {
  return ''
}

function getHeader(parser: ParsedMail) {
  const header = {
    from: '',
    to: '',
    cc: '',
    bcc: '',
    subject: '',
    date: '',
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
  const subject = parser.headers.get('subject')
  if (subject) {
    header.subject = (subject as string)
  }
  const date = parser.headers.get('date')
  if (date) {
    header.date = (date as Date).toLocaleString()
  }

  return header
}


function getHtml(parser: ParsedMail): string {
  let bodyStr = ''
  if (parser.html) {
    bodyStr = parser.html
  }
  const header = getHeader(parser)

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


for await (let message of mboxReader(readStream)) {
  console.log(message.returnPath);
  console.log(message.time);

  const parser = await simpleParser(message.content);
//  const mailBasename = getMailBaseName(parser)

  parser.attachments.forEach((attachment, index) => {
    console.log(attachment.filename)
    if (attachment.filename) {
      fs.writeFileSync(outputDir + '/' + attachment.filename, attachment.content);
    } else {
      fs.writeFileSync(outputDir + `/attach-${index}`, attachment.content);

    }
  })

  // console.log(`parser.html = ${parser.html}`)

  if (parser.html) {
    // console.log(parser.headers)
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    //set the HTML of this page
    await page.setContent(getHtml(parser));
    //save the page into a PDF and call it 'puppeteer-example.pdf'
    await page.pdf({ path: outputDir + `/body.pdf`, printBackground: true });
    //when, everything's done, close the browser instance.
    await browser.close();
    break
  }
}


console.log('DONE')


// TODO
// - eml is attached => no filename of the attachment, and the real attachment is in the eml that is attached
// - add link to attachment in the body.pdf
