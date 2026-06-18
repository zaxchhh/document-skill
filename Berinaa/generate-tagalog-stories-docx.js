const docx = require('docx');
const fs = require('fs');

const {
    Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, UnderlineType, PageBreak, TabStopType, TabStopPosition, Table, TableRow, TableCell, WidthType
} = docx;

async function generateDocx() {
    const doc = new Document({
        sections: [
            // Title Page
            {
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Exploring Short Tagalog Stories: Themes, Moral Lessons, and Cultural Significance",
                                size: 48,
                                bold: true,
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
                        children: [new TextRun({ text: "", break: 10 })],
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
                                        text: "",
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
                                text: "Short Tagalog stories, or \"maikling kuwento,\" are a cornerstone of Filipino literature, reflecting the nation's diverse culture, traditions, and values. These narratives often serve as moral compasses, teaching valuable lessons through relatable characters and scenarios.",
                                font: "Times New Roman",
                            }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Common Themes:",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        style: "Heading1",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Family and Kinship: Strong emphasis on familial bonds, respect for elders, and the challenges within family structures.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Love and Sacrifice: Romantic love, selfless acts for loved ones, and the pain of loss.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Poverty and Resilience: Stories depicting the struggles of the less fortunate and their unwavering spirit.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Superstition and Folklore: Integration of local beliefs, mythical creatures, and ancient traditions.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Nationalism and Identity: Narratives that explore what it means to be Filipino, often set against historical or social backdrops.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Moral Lessons:",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        style: "Heading1",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Short Tagalog stories are renowned for their embedded moral lessons, often presented explicitly at the end or subtly woven throughout the plot. These include:", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- The importance of honesty and integrity.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- The consequences of greed and selfishness.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- The value of hard work and perseverance.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- The power of forgiveness and reconciliation.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Respect for nature and the environment.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),

                    new Paragraph({
                        text: "Cultural Significance:",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        style: "Heading1",
                        run: { font: "Times New Roman" },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Preservation of Language: These stories play a vital role in keeping the Tagalog language vibrant and evolving.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Transmission of Values: They serve as a vehicle for passing down cultural values and ethics from one generation to the next.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Reflection of Society: They offer a mirror to Filipino society, showcasing its complexities, struggles, and aspirations.", font: "Times New Roman" }),
                        ],
                        spacing: { line: 360 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "- Educational Tool: Often used in schools to teach literature, values, and critical thinking.", font: "Times New Roman" }),
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
                                text: "Short Tagalog stories are more than just literary pieces; they are cultural artifacts that encapsulate the Filipino spirit. Their enduring popularity lies in their ability to entertain, educate, and resonate deeply with the experiences of the Filipino people, ensuring their place in the heart of the nation's literary heritage.",
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
    fs.writeFileSync("short-tagalog-stories-research.docx", Buffer.from(b64string, "base64"));
    console.log("Document generated successfully: short-tagalog-stories-research.docx");
}

generateDocx().catch((err) => {
    console.error("Error generating document:", err);
});