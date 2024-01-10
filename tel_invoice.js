const fs = require('fs');
const {
    Paragraph,
    patchDocument,
    PatchType,
    TextRun,
    AlignmentType
} = require('docx');
const path = require('path');
const config = require("./config.json");
const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);


const numRandomDatesPerMonth = config.receiptPerMonth;
const startYear = config.startYear;
const endYear = config.endYear;
const startMonth = config.startMonth;
const endMonth = config.endMonth
const prices = config.prices
const petrolPrices = [101.76, 101.21, 100.02, 102.98, 103.82];

const pertrolPumps = [
    {
        location: "Mahadevapura, Bengaluru -",
        pin: 560048
    },
    {
        location: "Immadihalli, Bengaluru -",
        pin: 560066
    },
    {
        location: "ITPL Main Rd, Bengaluru -",
        pin: 560037
    },
    {
        location: "KR Puram, Bengaluru -",
        pin: 560016
    },

]

function getPetrolBunk() {
    const randomIndex = Math.floor(Math.random() * pertrolPumps.length);
    // Retrieve the random item from the array
    const randomItem = pertrolPumps[randomIndex];

    return randomItem;
}


function getPriceAndLit() {
    const randomIndex = Math.floor(Math.random() * prices.length);
    // Retrieve the random item from the array
    const randomItem = prices[randomIndex];

    const randomPertrolIndex = Math.floor(Math.random() * petrolPrices.length);
    const randomPetrolItem = petrolPrices[randomIndex];

    const lit = randomItem / randomPetrolItem;

    return {
        lit: lit.toFixed(2),
        price: randomItem,
        petrolPrice: randomPetrolItem
    }

}

function getRandomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}


function formatMonth(month) {
    return new Intl.DateTimeFormat('en-US', { month: 'short' }).format(new Date(2000, month - 1, 1));
}

function getRandomDate(year, month) {
    const daysInMonth = new Date(year, month, 0).getDate();
    const day = getRandomInt(1, daysInMonth);

    const formattedDate = { date: `10 ${formatMonth(month)}`, year: `${year}` }
    return formattedDate;
}


function generateRandomDates(numPerMonth, startYear, endYear) {
    const randomDates = [];

    for (let year = startYear; year <= endYear; year++) {
        for (let month = startMonth; month <= endMonth; month++) {
            const uniqueDates = new Set();

            while (uniqueDates.size < numPerMonth) {
                const randomDate = getRandomDate(year, month);
                uniqueDates.add(randomDate);
            }

            randomDates.push(...Array.from(uniqueDates));
        }
    }

    return randomDates;
}


function generateRandomTime() {
    var hours = Math.floor(Math.random() * 24);
    var minutes = Math.floor(Math.random() * 60);
    var formattedHours = hours < 10 ? "0" + hours : hours;
    var formattedMinutes = minutes < 10 ? "0" + minutes : minutes;
    var randomTime = formattedHours + ":" + formattedMinutes;
    return randomTime;
}


function generateRandomNumber() {
    var useFiveChars = Math.random() < 0.5;

    var randomNumber = useFiveChars
        ? Math.floor(Math.random() * (999999 - 10000 + 1)) + 10000
        : Math.floor(Math.random() * (999999 - 1000 + 1)) + 1000;
    return randomNumber.toString();
}


function createFolderIfNotExists(folderPath) {
    if (!fs.existsSync(folderPath)) {
        // The folder doesn't exist, so create it
        fs.mkdirSync(folderPath, { recursive: true });
        console.log(`Folder created: ${folderPath}`);
    } else {
        console.log(`Folder already exists: ${folderPath}`);
    }
}


function writeFileSyncWithFolderCheck(filePath, data) {
    const folderPath = path.dirname(filePath);

    // Create the folder if it doesn't exist
    createFolderIfNotExists(folderPath);

    // Write the file
    fs.writeFileSync(filePath, data);
    console.log(`File written: ${filePath}`);
}




const editDocx = async () => {


    try {
        const randomDatesArray = generateRandomDates(numRandomDatesPerMonth, startYear, endYear);
        let result = {};
        for (const item of randomDatesArray) {
            const priceObj = getPriceAndLit();
            const petrolBunk = getPetrolBunk();

            const gst = (priceObj.price * 9 / 100).toFixed(2);
            const invAmt = (priceObj.price - 2 * gst).toFixed(2);
            result[item.date] = result[item.date] ? result[item.date] + priceObj.price : priceObj.price;
            const doc = await patchDocument(fs.readFileSync('/home/basavaraja/Downloads/junk/Internet_Invoice_Template_3.docx'), {
                patches: {
                    INV: {
                        type: PatchType.DOCUMENT,
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `Invoice No: IN${generateRandomNumber()}`,
                                        font: {
                                            name: "Arial MT",
                                            // size: "5px",

                                        },
                                        bold: false,
                                        scale: 100,
                                        size: "6pt",
                                        color: "4B4A4A",

                                    }),

                                ],
                            })
                        ],
                    },
                    Date: {
                        type: PatchType.DOCUMENT,
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `Date: ${item.date} ${item.year}`,
                                        font: {
                                            name: "Arial MT",
                                            // size: "5px",

                                        },
                                        bold: false,
                                        scale: 100,
                                        size: "6pt",
                                        color: "4B4A4A",
                                        alignment: AlignmentType.RIGHT
                                    }),

                                ],
                            }),
                            // new Paragraph({
                            //     children: [
                            //         new TextRun({
                            //             text: `   ${item.year}`,
                            //             font: {
                            //                 name: "Arial MT",
                            //                 // size: "5px",

                            //             },
                            //             bold: false,
                            //             scale: 190,
                            //             size: "6pt",
                            //             color: "1c1c1c",
                            //         }),

                            //     ],
                            // }),
                        ],
                    },
                    amt: {
                        type: PatchType.DOCUMENT,
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `    AMOUNT:   ₹   ${priceObj.price}`,
                                        font: {
                                            name: "Arial MT",
                                            // size: "5px",

                                        },
                                        bold: false,
                                        scale: 100,
                                        size: "6pt",
                                        color: "4B4A4A",
                                        

                                    }),

                                ],
                            }),

                        ],
                    },

                    totalamt: {
                        type: PatchType.DOCUMENT,
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `    ₹   ${priceObj.price}`,
                                        font: {
                                            name: "Arial MT",
                                            // size: "5px",

                                        },
                                        bold: false,
                                        scale: 100,
                                        size: "6pt",
                                        color: "4B4A4A",
                                        

                                    }),

                                ],
                            }),

                        ],
                    },
                    INVTOTAL: {
                        type: PatchType.DOCUMENT,
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `Invoice Total:   ₹`,
                                        font: {
                                            name: "Arial MT",
                                            // size: "5px",

                                        },
                                        bold: false,
                                        scale: 100,
                                        size: "6pt",
                                        color: "4B4A4A",
                                        

                                    }),
                                    new TextRun({
                                        text: ` ${priceObj.price}`,
                                        font: {
                                            name: "Arial MT",
                                            // size: "5px",

                                        },
                                        bold: true,
                                        scale: 100,
                                        size: "6pt",
                                        color: "090808",                                       

                                    })

                                ],
                            }),                           
                        ],
                    },
                    GST: {
                        type: PatchType.DOCUMENT,
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `         ₹ ${gst}`,
                                        font: {
                                            name: "Arial MT",
                                            // size: "5px",

                                        },
                                        bold: false,
                                        scale: 100,
                                        size: "6pt",
                                        color: "4B4A4A",

                                    }),

                                ],
                            })
                        ],
                    },
                    INVAMT: {
                        type: PatchType.DOCUMENT,
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `          ${invAmt}`,
                                        font: {
                                            name: "Arial MT",
                                            // size: "5px",

                                        },
                                        bold: false,
                                        scale: 100,
                                        size: "6pt",
                                        color: "4B4A4A",
                                        alignment: AlignmentType.CENTER
                                    }),

                                ],
                            })
                        ],
                    },




                    // receipt: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: ` Bill Account Number: ACT${generateRandomNumber()}`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 200,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         }),

                    //     ],
                    // },
                    // time: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `   Time:`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 190,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         }),
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `   ${generateRandomTime()}`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 190,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         }),

                    //     ],
                    // },

                    // location: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `       ${petrolBunk.location}`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 190,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         })
                    //     ],
                    // },
                    // pin: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `                     ${petrolBunk.pin}`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",
                    //                     },
                    //                     bold: false,
                    //                     scale: 190,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         })
                    //     ],
                    // },

                    // veh_no: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `    VEH   NO:   ${config.vehicleNumber}`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 155,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         }),

                    //     ],
                    // },

                    // cust_name: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `    CUSTOMER   NAME:`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 155,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         }),

                    //     ],
                    // },


                    // vol: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `    VOLUME`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 155,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),
                    //                 new TextRun({
                    //                     text: `(`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 255,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),
                    //                 new TextRun({
                    //                     text: `LTR.`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 155,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),
                    //                 new TextRun({
                    //                     text: `)`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 255,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),
                    //                 new TextRun({
                    //                     text: `:   ${priceObj.lit} `,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 155,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),
                    //                 new TextRun({
                    //                     text: `  lt`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 255,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         }),


                    //     ],
                    // },


                    // rateLit: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `    RATE/LTR:   ₹   ${priceObj.petrolPrice}`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 155,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         }),

                    //     ],
                    // },
                    // veh_type: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `    VEH   TYPE: `,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 155,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),
                    //                 new TextRun({
                    //                     text: `  Petrol`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 220,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         }),

                    //     ],
                    // },
                    // prod: {
                    //     type: PatchType.DOCUMENT,
                    //     children: [
                    //         new Paragraph({
                    //             children: [
                    //                 new TextRun({
                    //                     text: `    PRODUCT: `,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 155,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),
                    //                 new TextRun({
                    //                     text: `  Petrol`,
                    //                     font: {
                    //                         name: "Arial MT",
                    //                         // size: "5px",

                    //                     },
                    //                     bold: false,
                    //                     scale: 220,
                    //                     size: "4.5pt",
                    //                     color: "1c1c1c",

                    //                 }),

                    //             ],
                    //         }),

                    //     ],
                    // },



                },
            });

            // /home/basavaraja/Downloads/junk/Internet_Invoice_Template_3.docx
            const filePath = path.join("/home/basavaraja/Downloads/junk/",config.name, `Internet_Invoice-${item.date.replaceAll(" ", "-")}-${item.year}.pdf`);
            // const filePath = `/home/basavaraja/Downloads/junk/Internet_Invoice_Template_10.pdf`;
            let pdfBuf = await libre.convertAsync(doc, ".pdf", undefined);

            writeFileSyncWithFolderCheck(filePath, pdfBuf);

        }


        const sItem = Object.keys(result);
        const finResult = {};
        for (let index = 0; index < sItem.length; index++) {
            const element = sItem[index];
            const month = element.split(" ")[1];
            finResult[month] = finResult[month] ? finResult[month] + result[sItem[index]] : result[sItem[index]];
        }
        console.log("Result:", finResult);

    } catch (error) {
        console.error(`Error: ${error}`);
    }
};

editDocx()
    .then(() => {
        console.log('Document edited successfully.');
    })
    .catch((error) => {
        console.error(`Failed to edit document: ${error}`);
    });