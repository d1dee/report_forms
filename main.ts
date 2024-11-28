// Patch a document with patches

import { access, constants, mkdir } from "node:fs";
import { T_WriteToDocx, writeToDocx } from "./src/write_docx.ts";
import { parseXlsx, T_ParsedLearners } from "./src/xlsx.ts";

import { parseArgs } from "@std/cli";
import { join } from "@std/path";

/*  1. Get terminal arguments
    2. Check files exists
    3. Parse xlsx
        - if multiple sheets parser as multiple sheets
    4. Report errors
    5. Create folder
    6. save per class

*/
const cliAlias = {
    skip_errors: "S",
    help: "h",
    excel_file: "e",
    out_dir: "o",
    word_template: "t",
    "temp_dir": "T",
};

const cliArgs = parseArgs(Deno.args, {
    alias: cliAlias,
    string: ["excel_file", "out_dir", "word_template", "temp_dir"],
    boolean: ["help", "skip_errors"],
    default: {
        out_dir: Deno.cwd() + "/out",
    },
});

// check if excel_file and out_folder exist
const excel_file = cliArgs.excel_file;
const word_template = cliArgs.word_template;
const out_dir = cliArgs.out_dir;

// Check out dir
function init() {
    access(out_dir, constants.F_OK, (err) => {
        if (err && err.code === "ENOENT") mkdir(out_dir, { recursive: true }, () => init());
        else if (!err) {
            // Check read and write access in out dir
            access(out_dir, constants.R_OK | constants.W_OK, (err) => {
                if (err) {
                    throw new Error(
                        "Can not read or write in the provide directory. chown the directory or provide one with read and write permission.",
                    );
                }
            });

            // Check Excel file
            if (!excel_file) {
                throw new Error("Excel file was not provided in the commandline args");
            }
            access(excel_file, constants.R_OK, (err) => {
                if (err) throw new Error("Can not read the provided excel file.");
            });

            // Check template file
            if (!word_template) {
                throw new Error("Template file was not provided in the commandline args");
            }
            access(word_template, constants.R_OK, (err) => {
                if (err) throw new Error("Can not read the provided template file.");
            });
            // Call main to proceed...
            main([excel_file, word_template, out_dir] as const);
        } else throw err;
    });
}

function main([excelFile, wordTemplate, outDir]: [string, string, string]) {
    // Read excel data
    const excelData = Deno.readFileSync(excelFile);

    //Parse data to json and return object
    const [learnerJson, sheets] = parseXlsx(excelData, {
        skipErrors: cliArgs.skip_errors,
    });

    // Write to docs for each student
    const template = Deno.readFileSync(wordTemplate);

    function makeSubDirs(
        cb: T_WriteToDocx,
        args: [{ [k: string]: T_ParsedLearners }, Uint8Array, string],
    ) {
        function processSheets(i: number) {
            if (i >= sheets.length) {
                console.log("Done...");
                return;
            }

            const p = join(outDir, sheets[i]);
            access(p, constants.F_OK, (err) => {
                if (err && err.code === "ENOENT") {
                    mkdir(p, (err) => {
                        if (err) throw err;
                        cb(args[0][sheets[i]], args[1], join(args[2], sheets[i])).finally(() =>
                            processSheets(++i)
                        );
                    });
                } else if (!err) {
                    access(p, constants.W_OK | constants.R_OK, (err) => {
                        if (err) throw err;
                        cb(args[0][sheets[i]], args[1], join(args[2], sheets[i])).finally(() =>
                            processSheets(++i)
                        );
                    });
                } else throw err;
            });
        }
        processSheets(0);
    }

    makeSubDirs(writeToDocx, [learnerJson, template, outDir]);
}

init();
