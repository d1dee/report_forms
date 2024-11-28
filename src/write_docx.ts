import { patchDocument, PatchType, TextRun } from "npm:docx";

import { join } from "@std/path";
import { writeFileSync } from "node:fs";
import { T_ParsedLearners } from "./xlsx.ts";

export type T_WriteToDocx = (
    parsedLearner: T_ParsedLearners,
    templatePath: Uint8Array,
    outDir: string,
) => Promise<Array<void>>;

export const writeToDocx: T_WriteToDocx = function (sheetData, template, outDir) {
    const [rows] = sheetData;

    const promises = rows.map((row) =>
        patchDocument({
            outputType: "nodebuffer",
            data: template,
            patches: {
                adm_number: {
                    type: PatchType.PARAGRAPH,
                    children: [new TextRun({ text: String(row.adm_number) })],
                },

                students_name: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({ text: row.students_name.toUpperCase() }),
                    ],
                },

                date_entered: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({ text: row.date_entered.format("DD MMMM YYYY") }),
                    ],
                },
                date_left: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({ text: row.date_left.format("DD MMMM YYYY") }),
                    ],
                },
                form_entered: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({ text: row.form_entered }),
                    ],
                },
                form_left: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({ text: row.form_left }),
                    ],
                },
                date_of_birth: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({ text: row.date_of_birth.format("DD MMMM YYYY") }),
                    ],
                },
                academic_ability: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({ text: row.academic_ability.toLowerCase() }),
                    ],
                },
                industry: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({ text: row.industry.toLowerCase() }),
                    ],
                },
                conduct: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({ text: row.conduct.toLowerCase() }),
                    ],
                },
                activities: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({
                            text: row.activities?.split(",").map((v) => v.trim()).join(", ")
                                .toLowerCase(),
                        }),
                    ],
                },
            },
        }).then((doc) => writeFileSync(join(outDir, `${row.adm_number}.docx`), doc)).catch((err) =>
            console.log(err)
        )
    );

    return Promise.all(promises);
};
