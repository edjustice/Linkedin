# Metvy-Linkedin-Invite-Messaging-Tool
Metvy-Linkedin-Invite-Messaging-Tool
## Requirements

`node` >= `v4.0.0`

**Note**: If `node` and `npm` are not installed, Install them from [here](https://nodejs.org/en/download/).
## 1. Linkedin Auto Connection Sender : 

## Installation

Install this tool using `npm`:

```bash
$ npm install -g linkedin-auto-connect
```

It installs two binaries: `linkedin-auto-connect` and `lac` to your system path.

## Usage

Use it as follows using `lac` command:

```bash
$ lac -u <enter_your_linkedin_email>
Enter LinkedIn password: *****
```
## 2. Linkedin Auto Message Sender : 

## First installation
Make sure you have NodeJS and NPM installed on your computer, then run the `npm install` command to download the project dependencies.

## Use
 - Create an Excel file `Recipients.xlsx` containing a column` firstname`, `lastname`,` message` and `pj` (all other columns will be ignored)
 - Execute the command `npm start`, enter the time (in milliseconds) that the program will have to wait between each sending of message, then wait until the end of the program.
