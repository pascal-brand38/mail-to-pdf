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


function getHtml(parser: ParsedMail): string {
  let fromStr = ''
  let toStr = ''
  let ccStr = ''
  let bccStr = ''
  let body = ''
  let subjectStr = ''
  let dateStr = ''
  if (parser.headers) {
    console.log(parser.headers)
    const from = parser.headers.get('from')
    if (from) {
      fromStr = (from as AddressObject).text
    }
    const to = parser.headers.get('to')
    if (to) {
      toStr = (to as AddressObject).text
    }
    const cc = parser.headers.get('cc')
    if (cc) {
      ccStr = (cc as AddressObject).text
    }
    const bcc = parser.headers.get('bcc')
    if (bcc) {
      bccStr = (bcc as AddressObject).text
    }
    const subject = parser.headers.get('subject')
    if (subject) {
      subjectStr = (subject as string)
    }
    const date = parser.headers.get('date')
    if (date) {
      dateStr = (date as Date).toLocaleString()
    }

  }
  if (parser.html) {
    body = parser.html
  }

  let html = ''
  html += `<div style="background-color:lightgrey;">`
  html += '<div><br></div>'
  html += `<div><strong>From:</strong> ${escape(fromStr)}</div>`
  html += `<div><strong>To:</strong> ${escape(toStr)}</div>`
  html += `<div><strong>Cc:</strong> ${escape(ccStr)}</div>`
  html += `<div><strong>Bcc:</strong> ${escape(bccStr)}</div>`
  html += `<div><strong>Subject:</strong> ${escape(subjectStr)}</div>`
  html += `<div><strong>Date:</strong> ${escape(dateStr)}</div>`
  html += '<div><br></div>'
  html += `</div>`
  html += '<div><br></div>'
  html += body
  return html
}


for await (let message of mboxReader(readStream)) {
  console.log(message.returnPath);
  console.log(message.time);
  //process.stdout.write(message.content);
  // console.log(message.content)

  const str = await new Response(message.content).text();
  // console.log(str)

  const parser = await simpleParser(message.content);
  //console.log(parser)
  console.log('---------------------')
  // console.log(parser.from)
  // console.log(parser.date)
  // console.log(parser.subject)
  // console.log(parser.attachments)
  // console.log(parser.text)

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
