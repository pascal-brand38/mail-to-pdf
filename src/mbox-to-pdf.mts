#!/usr/bin/env node

// Copyright (c) Pascal Brand
// MIT License

import * as fs from 'fs'

// @ts-ignore
import { mboxReader } from 'mbox-reader'  // scan messages

import { simpleParser } from 'mailparser' // parse a single message
import puppeteer from "puppeteer"       // save html text as pdf

const mboxPath = 'C:/tmp/mbox-to-pdf/INBOX'
const outputDir = 'C:/tmp/mbox-to-pdf/output'


const readStream = fs.createReadStream(mboxPath);


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

  // console.log(parser.html)
  // if (true || parser.html) {
  //     const doc = new jsPDF();
  //     // doc.text(parser.html, 10, 10);
  //     const pdf = await doc.html("<h1>Hello, jsPDF!</h1>")
  //     //await pdf.save(outputDir + `/body.pdf`);
  // } else {
  //     //throw 'ERROR PASCAL'
  // }

  console.log(`parser.html = ${parser.html}`)

  if (parser.html) {
    console.log(parser.headers)
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    //set the HTML of this page
    await page.setContent(parser.html);
    //save the page into a PDF and call it 'puppeteer-example.pdf'
    await page.pdf({ path: outputDir + `/body.pdf` });
    //when, everything's done, close the browser instance.
    await browser.close();
    break
  }
}


console.log('DONE')


// TODO
// - eml is attached => no filename of the attachment, and the real attachment is in the eml that is attached
