#!/usr/bin/env node

// Copyright (c) Pascal Brand
// MIT License

import * as fs from 'fs'
import {simpleParser} from 'mailparser'

// @ts-ignore
import {mboxReader} from 'mbox-reader'

const mboxPath = 'C:/tmp/mbox-to-pdf/INBOX'

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

    const parser = await simpleParser(str);
    console.log(parser.text)

    // break
}


console.log('DONE')
