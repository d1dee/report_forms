// @ts-types="https://cdn.jsdelivr.net/npm/dayjs@1.11.13/plugin/customParseFormat.d.ts"
import customParseFormat from "https://esm.sh/dayjs/plugin/customParseFormat";
// @ts-types="https://cdn.jsdelivr.net/npm/dayjs@1.11.13/index.d.ts"
import dayJs from "https://esm.sh/dayjs";
// @ts-types="https://cdn.sheetjs.com/xlsx-0.20.3/package/types/index.d.ts"
import xlsx from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";
import { z } from "https://deno.land/x/zod/mod.ts";

export type T_ParsedLearners = [
    Array<z.infer<typeof schema>>,
    Array<{ error: string }>,
];

type T_ParseXlsx = (
    file: Uint8Array,
    { skipErrors }: { skipErrors: boolean },
) => [{ [k: string]: T_ParsedLearners }, Array<string>];

dayJs.extend(customParseFormat);

const schema = z.object({
    students_name: z.string().trim().refine(
        (v) => v.split(" ").length > 1,
        "Only one name was supplied.",
    ),
    adm_number: z.union([z.number(), z.string()]),
    date_entered: z.date().transform((v) => dayJs(v)),
    date_left: z.date().transform((v) => dayJs(v)),
    form_entered: z.string(),
    form_left: z.string(),
    date_of_birth: z.date().transform((v) => dayJs(v)),
    activities: z.union([z.string(), z.undefined()]),
    responsibilities: z.union([z.string(), z.undefined()]),
    academic_ability: z.string(),
    industry: z.string(),
    conduct: z.string(),
});

/**
 * @returns: An Object with sheets name as the key and an array of all the rows int hat sheets
 */
export const parseXlsx: T_ParseXlsx = function (fileData, opts) {
    const workbook = xlsx.read(fileData, { cellDates: true, dense: true });

    const jsonArray = [] as unknown as Array<[string, T_ParsedLearners]>;

    for (const sheetName of workbook.SheetNames) {
        const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { dateNF: "yyyy-mm-dd" });
        if (!data.length) continue;
        const passed = [], errored = [];
        for (const l of data) {
            if (typeof l !== "object" || !l) continue;
            const { success, data, error } = schema.safeParse(l);

            success ? passed.push(data) : errored.push({ ...l, error: error.message });
        }

        jsonArray.push([sheetName, [passed, errored]]);
    }

    if (!jsonArray.length) {
        throw new Error(
            `No data available in sheets; ${workbook.SheetNames.join(", ")}.`,
        );
    }
    return [Object.fromEntries(jsonArray), workbook.SheetNames];
};
