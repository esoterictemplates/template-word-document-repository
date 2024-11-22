import { readdirSync, readFileSync, writeFileSync } from "fs";
import xmlFormat from 'xml-formatter';

/**
 * Gets files from a directory recursively.
 * @param {string} directory The directory to get the files from.
 */
function getFiles(directory) {
    const output = [];

    const files = readdirSync(directory);
    files.forEach(file => {
        const path = `${directory}/${file}`;
        if (file.endsWith(".xml")) {
            output.push(path);
        } else {
            try {
                output.push(...getFiles(path));
            } catch (error) {
                // Ignore errors.
            }
        }
    });

    return output;
}

getFiles(".").forEach((file) => {
    const content = readFileSync(file);
    const formatted = xmlFormat(content.toString());
    writeFileSync(file, formatted);
});
