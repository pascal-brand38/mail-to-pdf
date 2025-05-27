# mail-to-pdf


<div align="center" style="font-size: 32px; font-weight: 700;">
<br>
Save emails in pdf
<br>
including the attachments.
</div>
<br>

[![NPM Version](https://img.shields.io/npm/v/mail-to-pdf.svg)](https://npmjs.com/package/mail-to-pdf)
[![NPM Downloads](https://img.shields.io/npm/dm/mail-to-pdf.svg)](https://npmjs.com/package/mail-to-pdf)
[![NPM Type Definitions](https://img.shields.io/npm/types/mail-to-pdf)](https://npmjs.com/package/mail-to-pdf)
[![NPM Last Update](https://img.shields.io/npm/last-update/mail-to-pdf)](https://npmjs.com/package/mail-to-pdf)

<br>

## Introduction

```mail-to-pdf``` is aimed at saving emails (as mbox files) in pdf, including the attachments.

Each email of a mbox file is saved in a directory, named by its
sending date followed by the object of the email.
The mail body is saved as a pdf, and its name is the sending
date followed by the object of the email.
Attachments are also saved, in their original format (jpeg, mp3, zip file,...)


## Installation and usage

```
npm install -g mail-to-pdf
mail-to-pdf --input <mbox-file | directory> --output-dir /tmp/mail-to-pdf
```

## Usage options

* ```--input <mbox | directory>```: save as pdf the email provided by ```--input``` options, which is either a mbox file, or a directory. In this latter case, all mbox files of the directory are processed.
By default, the emails in thunderbird are processed (limitation: only valid on windows)

* ```--output-dir <directory>```: output directory where all the emails are stored. Note that each email is saved in its own directory, along with its attachments.

* ```--no-parallel```: by default, parallel processing is performed to speed-up the process. Use this option to have sequential processing

* ```--force```: emails that were already saved as pdf are not re-processed by default. Use this option to force regenerating them

* ```--dryrun```: inputs are processed (to get statistics), but no file is generated (neither pdf nor attachments).

## Developer corner

```
git clone git@github.com:pascal-brand38/mail-to-pdf.git
cd mail-to-pdf
npm install
npm run build
node dist/mail-to-pdf.mjs --help
```
