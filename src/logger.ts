import { createWriteStream, WriteStream } from "node:fs";

let instance: ReturnType<typeof createLogger> | null = null;

export default function () {
    if (!instance) instance = createLogger();
    return instance;
}

function createLogger() {
    const fileStream: WriteStream = createWriteStream(`${Deno.cwd()}/log.txt`, {
        autoClose: true,
        flags: "a",
    });

    return {
        log(data: any) {
            const timeStamp = new Date().toDateString();
            fileStream.write(
                `${timeStamp}: - 
                ${JSON.stringify(data)}\n`,
            );
        },
    };
}
