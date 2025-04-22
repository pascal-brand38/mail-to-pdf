#!/usr/bin/env node

// Copyright (c) Pascal Brand
// MIT License

import * as fs from 'fs'
import {simpleParser} from 'mailparser'

// @ts-ignore
import {mboxReader} from 'mbox-reader'

const mboxPath = 'C:/tmp/mbox-to-pdf/INBOX'
const outputDir = 'C:/tmp/mbox-to-pdf/output'

// const mboxStr = fs.readFileSync(mboxPath, 'utf-8')
// const parser = await simpleParser(mboxStr);
// console.log(parser.text)

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
    console.log(parser.from)
    console.log(parser.date)
    console.log(parser.subject)
    console.log(parser.attachments)
    // console.log(parser.text)

    parser.attachments.forEach((attachment, index) => {
        console.log(attachment.filename)
        if (attachment.filename) {
            fs.writeFileSync(outputDir + '/' + attachment.filename, attachment.content);
        } else {
            fs.writeFileSync(outputDir + `/attach-${index}`, attachment.content);

        }
    })

//    break
}


console.log('DONE')


// TODO
// - eml is attached => no filename of the attachment, and the real attachment is in the eml that is attached
