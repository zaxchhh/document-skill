const docx = require('docx');
const fs = require('fs');

const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, UnderlineType, PageBreak, TabStopType, TabStopPosition, Table, TableRow, TableCell, WidthType } = docx;

async function generateDocx() {
    const doc = new Document({
        sections: [
            // Title Page
            {
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Philippine National Heroes: Contributions, Legacy, and Impact on Nation Building",
                                size: 48, // Larger font for title
                                bold: true,
                                font: "Times New Roman",
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                        spacing: {
                            after: 2000, // Space after title
                        },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "A Research Document",
                                size: 24,
                                font: "Times New Roman",
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                        spacing: {
                            after: 2000,
                        },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Date: June 18, 2026",
                                size: 24,
                                font: "Times New Roman",
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                        spacing: {
                            after: 2000,
                        },
                    }),
                    new Paragraph({
                        children: [new TextRun({ text: "", break: 10 })], // Add vertical space
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "By: AI Assistant",
                                size: 24,
                                font: "Times New Roman",
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                    }),
                    new Paragraph({
                        children: [new PageBreak()],
                    }),
                ],
            },
            // Main Content
            {
                properties: {
                    page: {
                        margin: {
                            top: docx.convertInchesToTwip(1),
                            right: docx.convertInchesToTwip(1),
                            bottom: docx.convertInchesToTwip(1),
                            left: docx.convertInchesToTwip(1),
                        },
                    },
                },
                headers: {
                    default: new docx.Header({
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: "", // No header text on content pages
                                    }),
                                ],
                            }),
                        ],
                    }),
                },
                footers: {
                    default: new docx.Footer({
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: docx.PageNumber.CURRENT,
                                        font: "Times New Roman",
                                        size: 24,
                                    }),
                                    new TextRun({
                                        text: " / ",
                                        font: "Times New Roman",
                                        size: 24,
                                    }),
                                    new TextRun({
                                        text: docx.PageNumber.TOTAL_PAGES,
                                        font: "Times New Roman",
                                        size: 24,
                                    }),
                                ],
                                alignment: AlignmentType.CENTER,
                            }),
                        ],
                    }),
                },
                children: [
                    new Paragraph({
                        text: "Introduction",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        style: "Heading1",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "The Philippines boasts a rich tapestry of heroes whose sacrifices and leadership shaped the nation's journey toward independence and democracy.",
                                font: "Times New Roman",
                            }),
                        ],
                        spacing: { line: 360 }, // 1.5 line spacing (360 twips per half-line)
                    }),

                    new Paragraph({
                        text: "Key Heroes:",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        style: "Heading1",
                        run: { font: "Times New Roman" },
                    }),

                    new Paragraph({
                        text: "José Rizal (1861-1896)",
                        heading: HeadingLevel.HEADING_2,
                        spacing: { before: 100, after: 100 },
                        style: "Heading2",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- The national hero, physician, and writer", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Author of \"Noli Me Tangere\" and \"El Filibusterismo\"", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Advocated for peaceful reform through education and writing", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Executed by Spanish colonial authorities in 1896", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Andrés Bonifacio (1863-1897)",
                        heading: HeadingLevel.HEADING_2,
                        spacing: { before: 100, after: 100 },
                        style: "Heading2",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Founder of the Katipunan revolutionary society", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- \"Father of the Philippine Revolution\"", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Led the Cry of Pugad Lawin in 1896", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Overthrown and executed by rival revolutionary leaders", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Emilio Aguinaldo (1869-1964)",
                        heading: HeadingLevel.HEADING_2,
                        spacing: { before: 100, after: 100 },
                        style: "Heading2",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- First president of the Philippines", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Led revolutionary forces against Spain and later the US", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Served as president during multiple periods", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Key figure in early Philippine governance", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Apolinario Mabini (1864-1903)",
                        heading: HeadingLevel.HEADING_2,
                        spacing: { before: 100, after: 100 },
                        style: "Heading2",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- \"The Sublime Ilocano,\" revolutionary leader and statesman", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Formulated the first Philippine constitution", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Served as prime minister and foreign minister", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Advocate for Philippine independence", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Juan Luna (1857-1899)",
                        heading: HeadingLevel.HEADING_2,
                        spacing: { before: 100, after: 100 },
                        style: "Heading2",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Famous painter and revolutionary", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Created the iconic \"Spoliarium\" painting", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Served as consul in Spain", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Contributed financially to the revolutionary cause", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Impact and Legacy:",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        style: "Heading1",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Inspired generations with their commitment to freedom and justice", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Established principles of nationalism and civic responsibility", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Their writings and actions continue to influence Philippine politics and culture", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Memorialized through holidays, monuments, and educational curricula", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Conclusion:",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        style: "Heading1",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Philippine heroes remain enduring symbols of courage, patriotism, and the enduring struggle for freedom. Their legacy continues to inspire Filipinos to strive for a better nation.",
                                font: "Times New Roman",
                            }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Bibliography:",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        style: "Heading1",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "[References would be added here]", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                ],
            },
        ],
    });

    const b64string = await Packer.toBase64String(doc);
    fs.writeFileSync("philippine-heroes-research.docx", Buffer.from(b64string, "base64"));
    console.log("Document generated successfully: philippine-heroes-research.docx");
}

generateDocx().catch((err) => {
    console.error("Error generating document:", err);
});
