import * as fs from "fs";
import { patchDocument, PatchType, TextRun } from "docx";

async function main() {

    const render = async (data: Uint8Array) => {
        const bufferData = Buffer.from(data.buffer, data.byteOffset, data.byteLength);

        return await patchDocument(bufferData, {
            keepOriginalStyles: true,
            patches: {
                p_x: {
                    type: PatchType.PARAGRAPH,
                    children: [new TextRun("Đầu máy ")]
                },
                gioi_tinh: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun({
                            text: "Ông",
                            bold: true,
                        })
                    ]
                },
                chuc_danh: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun("Lái đầu máy xe lửa")
                    ]
                },
                bac: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun("Bậc 3/3")
                    ]
                },
                hsl: {
                    type: PatchType.PARAGRAPH,
                    children: [
                        new TextRun("hệ số lương 1,94")
                    ]
                },
                ho_va_ten: {
                    type: PatchType.PARAGRAPH,
                    children: [new TextRun("Nguyễn Thế Lân.")],
                },
            },
        })
    }


    let buf = await render(fs.readFileSync("Quyết định.docx"))

    for (let i = 0; i < 10; i++) {
        buf = await render(buf)
    }

    fs.writeFileSync("My Document.docx", buf);
}

main()