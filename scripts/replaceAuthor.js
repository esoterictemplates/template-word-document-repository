import { readFileSync, writeFileSync } from "fs";
import os from "os";

const home = os.homedir();

const git = home + "/.gitconfig";

const content = readFileSync(git).toString();

const userRegex = /\[user\]\s+(\w+ ?= ?"?.+"?\n\s*)*name ?= ?"?(?<name>.+)"?/
const userName = content.match(userRegex)?.groups?.name;

if (userName === null) {
    console.log("Unable to find user name in git config.");
    process.exit(1);
}

const core = "./docProps/core.xml";
const info = readFileSync(core).toString();

const lastModifiedByRegex = /<cp:lastModifiedBy>(?<name>.+)<\/cp:lastModifiedBy>/;

const newInfo = info.replace(lastModifiedByRegex, `<cp:lastModifiedBy>${userName}</cp:lastModifiedBy>`);
writeFileSync(core, newInfo);
